import json, logging, subprocess, sys
from datetime import datetime, timedelta
from typing import Any, Callable, List, Optional

import requests
from django.utils import timezone
from playwright.sync_api import sync_playwright
from requests.auth import HTTPBasicAuth
from typing_extensions import TypedDict

from core.management.base_command import CoreBaseCommand
from core.models import SystemSanityCheck
from shadow_configs import alertmanager_config, rabbitmq_config
from shadow_helpers import execute_query, try_to_request

logger = logging.getLogger('jenkins_job')


class AlertmanagerLabels(TypedDict):
	alertname: str
	severity: str
	service: str


class AlertmanagerAnnotations(TypedDict):
	summary: str
	description: str


class AlertmanagerItem(TypedDict):
	labels: AlertmanagerLabels
	annotations: AlertmanagerAnnotations
	startsAt: str
	endsAt: Optional[str]


class HealthCheckNames(TypedDict):
	aws_rds: str
	rabbitmq: str
	react: str
	tms: str


class HealthCheckResult(TypedDict):
	aws_rds: Optional[AlertmanagerItem]
	rabbitmq: Optional[AlertmanagerItem]
	react: Optional[AlertmanagerItem]
	tms: Optional[AlertmanagerItem]
	# shadow_api: Optional[AlertmanagerItem]
	# meerkat: Optional[AlertmanagerItem]
	# monkey: Optional[AlertmanagerItem]
	# supervisor: Optional[AlertmanagerItem]
	# metachange: Optional[AlertmanagerItem]


class Command(CoreBaseCommand):
	_health_names: HealthCheckNames = {
		'aws_rds': 'AmazonRds',
		'rabbitmq': 'RabbitMq',
		'react': 'React',
		'tms': 'Tms',
	}

	def default_exception_handler(
		func: Callable[..., Optional[AlertmanagerItem]],
	) -> Callable[..., Optional[AlertmanagerItem]]:
		"""Return an alert even if an exception is raised

		This function is designed to be used in methods that will check
		a server and returns an alert if something goes wrong. Like the
		_check_* methods in this class.
		"""

		def wrapper(
			self, *args: Any, **kwargs: Any
		) -> Optional[AlertmanagerItem]:
			try:
				return func(self, *args, **kwargs)
			except Exception as e:
				return self._get_alert(
					func.__name__.replace('_check_', ''),
					'Exception',
					'An exception was raised while checking the health of the server',
					f'Exception: {e}',
				)

		return wrapper

	def __init__(self, *args: Any, **kwargs: Any) -> None:
		super().__init__(*args, **kwargs)

		self._health_result: HealthCheckResult = {
			'aws_rds': None,
			'rabbitmq': None,
			'react': None,
			'tms': None,
		}

		for server in self._health_result.keys():
			self._health_result[server] = self._get_alert(
				server,
				'CheckFailed',
				'Não foi possível realizar a verificação',
				'Por algum motivo não foi possível iniciar o processo de verificação do serviço',
			)

	def handle(self, *args: Any, **kwargs: Any) -> None:
		for server in self._health_result.keys():
			self._health_result[server] = getattr(self, f'_check_{server}')()

		try:
			alerts = [i for i in self._health_result.values() if i]
			self._send_alerts(alerts)
		except Exception as e:
			SystemSanityCheck.create(
				exception=e, job='run_general_health_check'
			)
			logger.critical(
				f'Error sending {self._health_names[server]} alert!'
			)

		logger.info(self._health_result)

	@default_exception_handler
	def _check_aws_rds(self) -> Optional[AlertmanagerItem]:
		query_result = execute_query('SELECT 1')
		if list(query_result[0].values())[0] != 1:
			return self._get_alert(
				'aws_rds',
				'NotWorking',
				'Banco de dados não está respondendo como esperado',
				'O banco de dados não está conseguindo responder um simples SELECT 1',
			)

	@default_exception_handler
	def _check_rabbitmq(self) -> Optional[AlertmanagerItem]:
		def has_download_stuck():
			query_result = execute_query(
				"""
				SELECT 1
				FROM shadowgiraffe.async_process ap
				INNER JOIN shadowgiraffe.async_process_status aps ON
					aps.id_status = ap.fk_async_process_status
				WHERE 1=1
				AND aps."name" = 'created'
				AND ap.created_at >= NOW() - INTERVAL '24 HOURS'
				AND ap.created_at <= NOW() - INTERVAL '30 MINUTES'
				ORDER BY ap.created_at DESC
			"""
			)
			return bool(len(query_result))

		# Read /api/index.html for more information about RabbitMQ API
		config = rabbitmq_config.config
		url = f"http://{config['host']}:15672/api/aliveness-test/%2F"
		response = requests.get(
			url, auth=HTTPBasicAuth(config['username'], config['password'])
		)
		if response.status_code == 200:
			data = response.json()
			if data['status'] == 'ok':
				if has_download_stuck():
					return self._get_alert(
						'rabbitmq',
						'DownloadStuck',
						'RabbitMQ está com um download na fila de espera por mais de 30 minutos',
						'Provavelmente tem algum download muito grande travando a fila do RabbitMQ',
					)
			else:
				return self._get_alert(
					'rabbitmq',
					'NotWorking',
					'RabbitMQ não está funcionando corretamente',
					f"O próprio health check na API do RabbitMQ /api/aliveness-test/%2F não está respondendo como esperado. O status retornado é '{data['status']}'",
				)
		else:
			return self._get_alert(
				'rabbitmq',
				'NotWorking',
				'RabbitMQ não está funcionando corretamente',
				f'Requisição feita para o RabbitMQ retornando um status HTTP diferente de 200. O status HTTP retornado é {response.status_code}',
			)

	@default_exception_handler
	def _check_react(self) -> Optional[AlertmanagerItem]:
		response = requests.get('https://app2.algumaempresa.com.br', timeout=60)
		if response.status_code == 200:
			if 'Erro de conexão' in self._fetch_html(
				'https://app2.algumaempresa.com.br'
			):
				return self._get_alert(
					'react',
					'ConnectionError',
					'React está apresentando um erro de conexão com o back-end',
					'Quando isso acontece, o usuário é redirecionado para a página de login, e aparece a mensagem "Erro de conexão" na tela. Provavelmente algum erro de conexão entre o front e o back-end está acontecendo',
				)
		else:
			return self._get_alert(
				'react',
				'NotWorking',
				'React não está funcionando corretamente',
				f'A página inicial da Alguma Empresa está respondendo com um status HTTP diferente de 200. O status HTTP retornado é {response.status_code}',
			)

	@default_exception_handler
	def _check_tms(self) -> Optional[AlertmanagerItem]:
		response = requests.get('https://tms.algumaempresa.com.br', timeout=60)
		if response.status_code != 200:
			return self._get_alert(
				'tms',
				'NotWorking',
				'TMS não está funcionando corretamente',
				f'O TMS está respondendo com um status HTTP diferente de 200. O status HTTP retornado é {response.status_code}',
			)

	@staticmethod
	def _fetch_html(url: str) -> str:
		with sync_playwright() as p:
			try:
				browser = p.firefox.launch(headless=True)
			except:
				subprocess.call(
					f'{sys.executable} -m playwright install firefox',
					shell=True,
				)
				browser = p.chromium.launch(headless=True)
			page = browser.new_page()
			page.goto(url)
			page.wait_for_load_state('networkidle')
			page.wait_for_selector('#page', timeout=30 * 1000)
			html = page.content()
			browser.close()
			return html

	@classmethod
	def _get_alert(
		cls, server: str, alertname: str, summary: str, description: str
	) -> AlertmanagerItem:
		now = timezone.make_aware(datetime.now())
		return {
			'labels': {
				'alertname': f'{cls._health_names[server]}{alertname}',
				'severity': 'critical',
				'server': server,
			},
			'annotations': {'summary': summary, 'description': description},
			'startsAt': now.isoformat(),
			'endsAt': (now + timedelta(minutes=20)).isoformat(),
		}

	@staticmethod
	def _send_alerts(alerts: List[AlertmanagerItem]) -> None:
		if alerts:
			config = alertmanager_config.config
			url = f"http://{config['server']}:{config['port']}/api/v2/alerts"
			try_to_request(
				method='POST',
				url=url,
				data=json.dumps(alerts),
				headers={'Content-Type': 'application/json'},
			)
