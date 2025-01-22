import os
import time
from enum import IntEnum
from typing import Iterator

import click
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    TimeoutException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

from autosig2.config import config
from autosig2.database import execute_query
from autosig2.logging import get_logger
from autosig2.utilities.dictionary import filter_dict_by_keys
from autosig2.utilities.files import find_and_remove_duplicates, remove_files_by_pattern
from autosig2.utilities.path import combine, makedirs
from autosig2.utilities.safety import safecall  # type: ignore[attr-defined]
from autosig2.utilities.string import is_substring_normalized, normalized_string_comparison
from autosig2.utilities.webdriver import BaseWebsite, WebsiteException
from autosig2.utilities.workbook import get_table_rows_from_workbook
from autosig2.wrappers import wclick

logger = get_logger(__name__)


def get_logins() -> Iterator[tuple[str, str, str, str, str]]:
    corretoras = [
        c['dsnfantasia']
        for c in execute_query("SELECT DISTINCT dsnfantasia FROM corretora")
    ]
    estipulantes = [
        e['dsnfantasia']
        for e in execute_query("SELECT DISTINCT dsnfantasia FROM estipulante")
    ]

    def get_corretora(row: dict[str, str]) -> str:
        for corretora in corretoras:
            if is_substring_normalized(corretora, row["corretora"]):
                return corretora.strip()
        return row["corretora"]

    def get_estipulante(row: dict[str, str]) -> str:
        for estipulante in estipulantes:
            if is_substring_normalized(estipulante, row["estipulante"]):
                return estipulante.strip()
        return row["estipulante"]

    workbook_path = combine(
        config["network"]["swap_path"], r".\Pass\Alguma Empresa\Acessos Saúde Online.xlsx"
    )
    for row in get_table_rows_from_workbook(workbook_path):
        if filter_dict_by_keys(row, ["corretora", "estipulante", "login", "usuario", "senha"]):
            yield get_corretora(row), get_estipulante(row), row["login"], row["usuario"], row["senha"]
        else:
            raise ValueError("Invalid row in workbook")


def wait_for_downloads(download_path: str, timeout: int = 60) -> bool:
    end_time = time.time() + timeout
    while time.time() < end_time:
        files_in_progress = [
            filename
            for filename in os.listdir(download_path)
            if (
                filename.endswith(".crdownload")
                or filename.endswith(".part")
                or filename.endswith(".tmp")
            )
        ]
        if not files_in_progress:
            return True
        time.sleep(1)
    return False


class AlgumaEmpresaPortalDownloadType(IntEnum):
    PDF = 1
    TXT = 2
    CSV = 3


class AlgumaEmpresaPortal(BaseWebsite):
    def __init__(
        self,
        code: str,
        user: str,
        password: str,
        download_path: str,
        download_type: AlgumaEmpresaPortalDownloadType,
        is_headless: bool,
    ) -> None:
        super().__init__(
            download_path=download_path,
            insecure_origins_ok=["http://algumaempresa.com.br"],
            is_headless=is_headless,
        )
        self.code = code
        self.user = user
        self.password = password
        self.download_type = download_type

    def run(self) -> None:
        self._login()
        self._premio()
        self._gerencial()

    def _login(self) -> None:
        self.get_page("https://algumaempresaseguros.com.br/empresa/login/")
        self.perform_actions(
            {
                "locators": {
                    "code": (By.ID, "code"),
                    "user": (By.ID, "user"),
                    "password": (By.ID, "senha"),
                    "login": (By.ID, "entrarLogin"),
                },
                "actions": [
                    ("send_keys", "code", self.code),
                    ("send_keys", "user", self.user),
                    ("send_keys", "password", self.password),
                    ("click", "login"),
                ],
            }
        )

    def _premio(self) -> None:
        def download_from_url(url: str) -> None:
            self.get_page(url)
            download_start = 2
            download_end = self.count_elements((By.CSS_SELECTOR, ".tablist tr"))
            while True:
                try:
                    self.perform_actions(
                        {
                            "locators": {
                                f"{i}": (
                                    By.XPATH,
                                    f"//tr[{i}]/td[3]/a[{int(self.download_type)}]/img",
                                )
                                for i in range(download_start, download_end + 1)
                            },
                            "actions": [
                                ("click", f"{i}") for i in range(download_start, download_end + 1)
                            ],
                        },
                        delay=5,
                    )
                    break
                except WebsiteException as exception:
                    locator = exception.context["locator"]
                    download_start = int(locator)
                    if type(exception.raised_exception) == UnexpectedAlertPresentException:
                        try:
                            # NOTE: Maybe can be interesting to increase the timeout
                            # to e. g. 15 seconds.
                            custom_waiter = WebDriverWait(self.driver, timeout=10)
                            alert = custom_waiter.until(EC.alert_is_present())
                            alert.accept()
                        except TimeoutException:
                            pass
                    else:
                        raise exception.raised_exception

        download_from_url(
            "https://algumaempresa.com.br/empresa/faturamento/emissao-de-fatura-de-premio/anexos/movimentacoes-ocorridas-revitalizado.htm"
        )
        download_from_url(
            "https://algumaempresa.com.br/empresa/faturamento/emissao-de-fatura-de-premio/anexos/emissao-de-fatura-grupal-revitalizado.htm"
        )

    def _gerencial(self) -> None:
        self.get_page(
            "https://algumaempresa.com.br/empresa/informacoes-gerenciais/relatorios-tecnicos/"
        )
        try:
            self.perform_actions(
                {
                    "locators": {
                        # "proceed": (By.ID, "proceed-button"),
                        "sinistralidade": (By.LINK_TEXT, "Sinistralidade Pagamento"),
                        # TODO(cahian): Add support for custom download types
                        "download": (By.ID, "exportPDF:exportIcon"),  # exportPDF:exportIconCSV
                    },
                    "actions": [
                        # ("click", "proceed"),
                        ("click", "sinistralidade"),
                        ("click", "download"),
                    ],
                }
            )
        except WebsiteException as exception:
            error_message = "Informamos que devido às características deste produto, não geramos relatório gerencial para essa empresa/apólice."
            if (
                type(exception.raised_exception) == ElementClickInterceptedException
                and error_message in self.driver.page_source
            ):
                logger.info("Relatório gerencial não suportado nesse acesso.")
            else:
                raise exception.raised_exception


@wclick.command()
@click.option("-c", "--corretora", "corretoras", multiple=True, help="Especifica corretora")
@click.option("-e", "--estipulante", "estipulantes", multiple=True, help="Especifica estipulante")
@click.option("-h", "--headless", is_flag=True, help="Roda em modo headless")
@click.option("-f", "--filetype", type=click.Choice(["pdf", "txt", "csv"]), default="csv")
def obter(corretoras: tuple[str], estipulantes: tuple[str], headless: bool, filetype: str) -> None:
    """Obter os relatórios de prêmio e gerencial de Alguma Empresa"""
    for corretora, estipulante, code, user, password in get_logins():
        if corretoras and not normalized_string_comparison(corretora, corretoras):
            continue
        if estipulantes and not normalized_string_comparison(estipulante, estipulantes):
            continue
        logger.info(
            f'Obtendo relatórios de Alguma Empresa para a estipulante "{estipulante}" no acesso ({code}, {user}, {password}).'
        )
        download_path = makedirs(
            config["network"]["integration_path"],
            r"I:\recepcao\Alguma Empresa\Automatização\Website",
            corretora.title(),
            estipulante.title(),
        )
        download_type = AlgumaEmpresaPortalDownloadType[filetype.upper()]

        try:
            portal = AlgumaEmpresaPortal(code, user, password, download_path, download_type, headless)
            safecall(portal.run)
            wait_for_downloads(download_path)
            portal.close()
        except Exception as e:
            portal.close()
            raise e
        finally:
            try:
                portal.close()
            except:
                pass

        # TODO: Rename all sinistralidadepagamento*.pdf to add the login code
        # For example, read the sinistralidadepagamento, get the code, and rename
        # from "sinistralidadepagamento (1).pdf" for "sinistralidadepagamento - {code}.pdf"
        find_and_remove_duplicates(download_path)
        remove_files_by_pattern(download_path, "*.tmp")
        remove_files_by_pattern(download_path, "*.crdownload")
        remove_files_by_pattern(download_path, "*.part")
