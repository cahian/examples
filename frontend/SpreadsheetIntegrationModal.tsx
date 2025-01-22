import React, {
	useState,
	useEffect,
	useMemo,
	useCallback,
} from 'react'
import { MdLightbulb } from 'react-icons/md'
import Swal from 'sweetalert2'
import {
	Button,
	Modal,
	ModalStepper,
	LoadingAndError,
} from 'components'
import FormInputs, {
	FormValues,
} from 'components/FormInputs/FormInputs/FormInputs'
import {
	IFormInput,
	InputType,
} from 'components/FormInputs/FormInput'
import { NestedInputProps } from 'components/Input/NestedInput/NestedInput'
import algumaEmpresaAPI from 'services/axios/axios'
import { store } from 'services/redux'
import errorSwal from 'utils/errorSwal/errorSwal'
import { SpreadsheetIntegrationType } from 'views/IntegrationPanel/Types'
import Switch from '../../../components/Input/Switch/Switch'

import './SpreadsheetIntegrationModal.scss'

interface SpreadsheetTutorialModalProps {
	isOpen: boolean
	setIsOpen: (open: boolean) => void
}

const SpreadsheetTutorialModal: React.FC<
	SpreadsheetTutorialModalProps
> = ({ isOpen, setIsOpen }) => {
	return (
		<Modal
			open={isOpen}
			onClose={() => setIsOpen(false)}
			className='modal-stepper-spreadsheet-tutorial'
			modalType='center'
			size='lg'
			title='Quando usar os códigos de produto e pedido?'
			hasCloseButton
		>
			<>
				<div className='tutorial-content'>
					<div className='guide-container'>
						Saiba quando é melhor incluir ou não os códigos de
						produto e pedido para o sistema.
					</div>
					<div className='modal-flex-container'>
						<div className='column'>
							<div className='column-title'>
								<Switch
									labelPlacement='start'
									defaultChecked={false}
									disabled
								/>
								Quando não usar código
							</div>
							<ul>
								<li>Ideal para preenchimento manual</li>
								<li>Simples para planilhas pequenas</li>
								<li>Quando há poucos pedidos e produtos</li>
							</ul>
						</div>
						<div className='column'>
							<div className='column-title'>
								<Switch
									labelPlacement='start'
									defaultChecked
									disabled
								/>
								Quando usar código
							</div>
							<ul>
								<li>Melhor se você já tem os códigos</li>
								<li>Útil para integrações com bases existentes</li>
								<li>Facilita a organização e rastreamento</li>
							</ul>
						</div>
					</div>
				</div>
				<div className='button-content'>
					<Button onClick={() => setIsOpen(false)}>
						Voltar
					</Button>
				</div>
			</>
		</Modal>
	)
}

type CompanyCrudResultType = {
	company_name: string
	has_catalog: boolean
}

interface SpreadsheetFormProps {
	formValues: FormValues
	handleChange: (inputName: string, inputValue: any) => void
}

const SpreadsheetForm: React.FC<SpreadsheetFormProps> = ({
	formValues,
	handleChange,
}) => {
	const { companyName } = store.getState().company
	const [companyHasCatalog, setCompanyHasCatalog] =
		useState<boolean>()
	const [
		companyHasCatalogPromise,
		setCompanyHasCatalogPromise,
	] = useState<Promise<any>>()

	useEffect(() => {
		setCompanyHasCatalogPromise(
			algumaEmpresaAPI({
				url: 'crud/company/list',
				method: 'POST',
			}).then((response: any) => {
				let company: CompanyCrudResultType | null = null
				response.data.results.forEach(
					(companyResult: CompanyCrudResultType) => {
						if (companyResult.company_name === companyName) {
							if (company) {
								throw new Error(
									'Mais de uma company com o mesmo nome encontrada',
								)
							}
							company = companyResult
						}
					},
				)

				if (company) {
					if ('has_catalog' in company) {
						setCompanyHasCatalog(
							(company as { has_catalog: boolean }).has_catalog,
						)
					} else {
						throw new Error('Atributo de catalogo não encontrado')
					}
				} else {
					throw new Error('Nenhuma company encontrada')
				}
			}),
		)
	}, [companyName])

	const formData = useMemo(() => {
		let data: IFormInput[] = [
			{
				type: InputType.switch,
				name: 'has_product_code',
				label: 'Desejo enviar o código de produto',
				required: true,
				initial_data: formValues.has_product_code,
			},
			{
				type: InputType.switch,
				name: 'has_order_code',
				label: 'Desejo enviar o código de pedido',
				required: true,
				initial_data: formValues.has_order_code,
			},
		]
		if (!companyHasCatalog) {
			const item = data.find(
				(i) => i.name === 'has_order_code',
			)
			if (item) {
				data = [item]
			} else {
				throw new Error('Company has_order_code not found')
			}
		}
		return data
	}, [formValues, companyHasCatalog])

	return (
		<LoadingAndError promise={companyHasCatalogPromise}>
			<div className='spreadsheet-form'>
				<FormInputs
					formData={formData}
					onChange={handleChange}
				/>
			</div>
		</LoadingAndError>
	)
}

const nestedInputFunction = (formValues: FormValues) => {
	if (formValues.websites) {
		const oldFormWebsites =
			typeof formValues.websites === 'string'
				? JSON.parse(formValues.websites)
				: formValues.websites
		const newFormWebsites: NestedInputProps[][] =
			oldFormWebsites.map(
				(website: {
					website_name: string
					is_online_channel: boolean
					is_marketplace_out: boolean
				}) => {
					const nestedInputs: NestedInputProps[] = [
						{
							type: InputType.text,
							name: 'website_name',
							label: 'Nome do canal de venda',
							required: true,
							initial_data: website.website_name,
						},
						{
							type: InputType.switch,
							name: 'is_online_channel',
							label: 'É venda online?',
							required: true,
							initial_data: website.is_online_channel,
						},
						{
							type: InputType.switch,
							name: 'is_marketplace_out',
							label: 'É um Marketplace?',
							required: true,
							initial_data: website.is_marketplace_out,
						},
					]
					return nestedInputs
				},
			)
		return newFormWebsites
	}

	const nestedInputs: NestedInputProps[][] = [
		[
			{
				type: InputType.text,
				name: 'website_name',
				label: 'Nome do canal de venda',
				required: true,
			},
			{
				type: InputType.switch,
				name: 'is_online_channel',
				label: 'É venda online?',
				required: true,
			},
			{
				type: InputType.switch,
				name: 'is_marketplace_out',
				label: 'É um Marketplace?',
				required: true,
			},
		],
	]
	return nestedInputs
}

interface SpreadsheetWebsitesFormProps {
	formValues: FormValues
	handleChange: (inputName: string, inputValue: any) => void
}

const SpreadsheetWebsitesForm: React.FC<
	SpreadsheetWebsitesFormProps
> = ({ formValues, handleChange }) => {
	const formWebsites: IFormInput[] = [
		{
			type: InputType.nested,
			name: 'websites',
			label: 'Canais de Venda',
			nestedInputs: nestedInputFunction(formValues),
		},
	]
	return (
		<div className='spreadsheet-form'>
			<FormInputs
				formData={formWebsites}
				onChange={handleChange}
			/>
		</div>
	)
}

interface SpreadsheetIntegrationModalProps {
	formValues: FormValues
	handleChange: (inputName: string, inputValue: any) => void
	spreadsheetConfig: SpreadsheetIntegrationType
	isOpen: boolean
	setIsOpen: (open: boolean) => void

	onSaveStart?: () => void
	onSaveError?: () => void
	onSaveSuccess?: () => void
	onSaveFinish?: () => void
}

const SpreadsheetIntegrationModal: React.FC<
	SpreadsheetIntegrationModalProps
> = ({
	formValues,
	handleChange,
	spreadsheetConfig,
	isOpen,
	setIsOpen,

	onSaveStart,
	onSaveError,
	onSaveSuccess,
	onSaveFinish,
}) => {
	const [isTutorialOpen, setIsTutorialOpen] =
		useState<boolean>(false)
	const [isError, setIsError] = useState<boolean>(false)

	const save: () => void = useCallback(() => {
		const url = `save_spreadsheet_integration`
		onSaveStart?.()

		const websiteNames: string[] = []
		if (formValues.websites) {
			for (let i = 0; i < formValues.websites.length; i += 1) {
				const website = formValues.websites[i]
				if (!websiteNames.includes(website.website_name)) {
					websiteNames.push(website.website_name)
				} else {
					onSaveError?.()
					errorSwal(
						'Não pode ter o mesmo canal de venda mais de uma vez',
					)
					return Promise.reject()
				}
				if (!website.website_name) {
					onSaveError?.()
					errorSwal(
						'Preencha o nome de todos os canais de venda',
					)
					return Promise.reject()
				}
			}
		}

		const promiseRequest = algumaEmpresaAPI({
			url: url,
			method: 'POST',
			data: formValues,
		})

		if (promiseRequest) {
			promiseRequest
				.then(() => {
					Swal.fire({
						title: 'Acesso à integração concedido com sucesso',
						icon: 'success',
						text: 'Continue com o preenchimento dos dados.',
					})
					onSaveSuccess?.()
					return true
				})
				.catch((error) => {
					const errors = error.response.data
					onSaveError?.()
					errorSwal(errors.errors)
					return false
				})
		} else {
			onSaveError?.()
			errorSwal('Dê um nome a todos os seus canais de venda')
		}

		return promiseRequest
	}, [formValues, onSaveStart, onSaveError, onSaveSuccess])

	const modalTitle = useMemo(
		() => spreadsheetConfig.title,
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[],
	)

	const stepsConfig = useMemo(
		() => [
			<div className='spreadsheet-integration-step'>
				<div
					key='integration-help'
					className='integration-help'
				>
					<span>
						<MdLightbulb />
					</span>
					<span onClick={() => setIsTutorialOpen(true)}>
						Ajuda com os códigos de produto e pedido
					</span>
				</div>

				<div className='guide-container'>
					Preencha as informações abaixo sobre suas preferências
					em relação ao envio das planilhas de integração.
				</div>

				<div className='align-center'>
					<SpreadsheetForm
						formValues={formValues}
						handleChange={handleChange}
					/>
				</div>

				<SpreadsheetTutorialModal
					isOpen={isTutorialOpen}
					setIsOpen={setIsTutorialOpen}
				/>
			</div>,
			<div className='spreadsheet-integration-step'>
				<div className='guide-container'>
					Preencha as informações abaixo sobre os seus canais de
					venda (você pode adicionar mais de um).
				</div>

				<div className='align-center'>
					<SpreadsheetWebsitesForm
						formValues={formValues}
						handleChange={handleChange}
					/>
				</div>
			</div>,
		],
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[formValues, isTutorialOpen],
	)

	useEffect(() => {
		let isMount = true

		algumaEmpresaAPI({
			url: 'get_configured_integrations',
			method: 'GET',
		})
			.then((response) => {
				if (isOpen && isMount) {
					const integrations: Array<{
						integration: string
						is_marketing_integration: boolean
					}> = response.data
					if (integrations) {
						const spreadsheetIntegrations = integrations.filter(
							(integ) => integ.integration === 'spreadsheet',
						)
						const nonSpreadsheetIntegrations =
							integrations.filter(
								(integ) =>
									integ.integration !== 'spreadsheet' &&
									!integ.is_marketing_integration,
							)
						if (spreadsheetIntegrations.length > 0) {
							return Promise.reject(
								new Error(
									'Não é possível configurar uma integração de planilha mais de uma vez',
								),
							)
						}
						if (nonSpreadsheetIntegrations.length > 0) {
							return Promise.reject(
								new Error(
									'Não é possível configurar uma integração de planilha uma vez que já existe uma integração de plataforma',
								),
							)
						}
					}
				}
				return response
			})
			.then(() => {
				setIsError(false)
			})
			.catch((error) => {
				setIsError(true)
				if (error?.response) {
					errorSwal(error.response.data.errors)
				} else if (error?.message) {
					errorSwal(error.message)
				}
			})

		return () => {
			isMount = false
		}
	}, [isOpen])

	return (
		<div className='spreadsheet-integration-modal-container'>
			<ModalStepper
				className='modal-stepper-spreadsheet'
				modalTitle={modalTitle || ''}
				stepsConfig={stepsConfig}
				isOpen={isOpen && !isError}
				setIsOpen={setIsOpen}
				save={save}
				onSave={onSaveFinish}
			/>
		</div>
	)
}

SpreadsheetIntegrationModal.defaultProps = {
	onSaveStart: undefined,
	onSaveError: undefined,
	onSaveSuccess: undefined,
	onSaveFinish: undefined,
}

export default SpreadsheetIntegrationModal
