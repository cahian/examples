/**
 * Smoke Tests
 *
 * Smoke tests are a set of basic, high-level tests that check whether the most
 * crucial functionality of an application works as expected. They are intended
 * to identify critical issues early, ensuring that the application is stable
 * enough for further testing.
 *
 * Smoke tests typically include verifying the following:
 * - Core application functionality (e.g., page loading, key user actions)
 * - Essential UI elements are rendered correctly
 * - Basic navigation works
 * - No significant errors are thrown
 *
 * These tests should be fast and should not cover edge cases or detailed
 * functionality.
 */

/// <reference types="cypress" />
import 'cypress-network-idle'
import { format as formatDate } from 'date-fns'

type LoginType = {
	username: string
	password: string
	companyHumanizedName: string
}

type LoginsType = {
	[key: string]: LoginType
}

type CapturedErrorType = {
	method: string
	url: string
	statusCode: number
	statusText: string
	message: string
}

type CapturedErrorsType = {
	[key: string]: CapturedErrorType[]
}

const DEFAULT_TIMEOUT = 100000

// The SIDEBAR_PAGES_TO_INCLUDE and SIDEBAR_PAGES_TO_EXCLUDE arrays
// control the sidebar page inclusion/exclusion:
// - If both are empty, all pages are included.
// - If SIDEBAR_PAGES_TO_INCLUDE has values, only those pages are included
// (excluding others, even if they are not in SIDEBAR_PAGES_TO_EXCLUDE).
// - If SIDEBAR_PAGES_TO_EXCLUDE has values, all pages except those are included
// (regardless of whether they are in SIDEBAR_PAGES_TO_INCLUDE).
// - If both have values, only pages in SIDEBAR_PAGES_TO_INCLUDE and not in
// SIDEBAR_PAGES_TO_EXCLUDE will be included.
const SIDEBAR_PAGES_TO_INCLUDE = []
const SIDEBAR_PAGES_TO_EXCLUDE = []

const SIDEBAR_PAGES_CUSTOM_ACTIONS = {
	daily_cashflow: () => {
		cy.contains('button', 'Confirmar').click()
	},
}

const logins: LoginsType = Cypress.env('logins')

const resetTerminalLogs = () => {
	Cypress.env('terminalLogs', [])
}

const addTerminalLog = (message: string) => {
	console.log(message)
	const terminalLogs = Cypress.env('terminalLogs') || []
	terminalLogs.push(message)
	Cypress.env('terminalLogs', terminalLogs)
}

const getLastClickedPageFromTerimnalLogs = () => {
	const terminalLogs = Cypress.env('terminalLogs') || []
	for (let i = terminalLogs.length - 1; i >= 0; i--) {
		if (terminalLogs[i].includes('clicked on')) {
			return terminalLogs[i].split('clicked on: ')[1]
		}
	}
	return null
}

const logAllTerminalLogs = () => {
	const terminalLogs = Cypress.env('terminalLogs') || []
	terminalLogs.forEach((log) => cy.task('log', log))
}

const addCapturedError = (
	testTitle: string,
	capturedError: CapturedErrorType,
) => {
	const capturedErrors: CapturedErrorsType =
		Cypress.env('capturedErrors') || {}
	if (!(testTitle in capturedErrors)) {
		capturedErrors[testTitle] = []
	}
	capturedErrors[testTitle].push(capturedError)
	Cypress.env('capturedErrors', capturedErrors)
}

const getCapturedErrors = () => {
	const capturedErrors: CapturedErrorsType =
		Cypress.env('capturedErrors') || {}
	return capturedErrors
}

const getTotalCapturedErrors = (
	capturedErrors: CapturedErrorsType,
) => {
	return Object.values(capturedErrors).reduce(
		(acc, errors) => acc + errors.length,
		0,
	)
}

const waitForNetworkIdlePrepare = () => {
	cy.waitForNetworkIdlePrepare({
		method: '*',
		pattern: '*',
		alias: 'calls',
	})
}

const waitForNetworkIdle = () => {
	cy.waitForNetworkIdle('@calls', 8000, {
		timeout: DEFAULT_TIMEOUT,
	})
}

const goToHomePage = (
	companyName: string,
	login: LoginType,
) => {
	cy
		.intercept('**/api/**', (req) => {
			req.continue((res) => {
				// Ignore login requests
				if (!req.url.includes('api/login')) {
					const pathname = new URL(req.url).pathname
					addTerminalLog(
						`execution - test ${companyName} - interpected the: ${pathname}`,
					)

					if (res.statusCode >= 400) {
						addTerminalLog(
							`execution - test ${companyName} - error on: ${pathname}`,
						)

						let message = `Error: Received status code ${res.statusCode} for ${req.url}`
						const lastClickedPage =
							getLastClickedPageFromTerimnalLogs()
						if (lastClickedPage) {
							message += ` on page ${lastClickedPage}`
						}

						addCapturedError(Cypress.currentTest.title, {
							method: req.method,
							url: req.url.trim(),
							statusCode: res.statusCode,
							statusText: res.statusMessage,
							message: message,
						})
					}
				}
			})
		})
		.as('algumaEmpresaApiRequests')

	waitForNetworkIdlePrepare()
	cy.visit(`/${companyName}`, { failOnStatusCode: false })
	waitForNetworkIdle()

	cy.get('#username').type(login.username)
	cy.get('#password').type(login.password)

	waitForNetworkIdlePrepare()
	cy.get('button[type=submit]').click()
	waitForNetworkIdle()
}

const getMailSubjectTemplate = (
	totalErrors: number,
	companyHumanizedName?: string,
) => {
	let subject = `Cypress Test Report`
	if (companyHumanizedName) {
		subject += ` - ${companyHumanizedName}`
	}
	subject += ` - ${totalErrors} Error${totalErrors > 1 ? 's' : ''} Captured`
	subject += ` - ${formatDate(new Date(), "yyyy-MM-dd'T'HH:mm:ss")}`
	return subject
}

const getMailHtmlTemplate = (
	capturedErrors: CapturedErrorsType,
	totalErrors: number,
) => {
	return `
		<!DOCTYPE html>
		<html lang="en">
		<head>
			<meta charset="UTF-8">
			<meta name="viewport" content="width=device-width, initial-scale=1.0">
			<title>Cypress Test Report</title>
			<style>
				body {
					font-family: Arial, sans-serif;
					line-height: 1.6;
				}
				h1 {
					color: #444;
				}
				h2 {
					color: #555;
					margin-top: 20px;
				}
				table {
					width: 100%;
					border-collapse: collapse;
					margin: 20px 0;
				}
				th, td {
					border: 1px solid #ddd;
					padding: 8px;
					text-align: left;
				}
				th {
					background-color: #f4f4f4;
				}
				tr:nth-child(even) {
					background-color: #f9f9f9;
				}
				tr:hover {
					background-color: #f1f1f1;
				}
			</style>
		</head>
		<body>
			<h1>Cypress Test Report</h1>
			<p>Total errors captured: <strong>${totalErrors}</strong></p>
			${Object.entries(capturedErrors)
				.map(
					([testName, errors]) => `
					<h2>${testName}</h2>
					<table>
						<thead>
							<tr>
								<th>#</th>
								<th>Method</th>
								<th>URL</th>
								<th>Status Code</th>
								<th>Status Text</th>
								<th>Message</th>
							</tr>
						</thead>
						<tbody>
							${errors
								.map(
									(error, index) => `
								<tr>
									<td>${index + 1}</td>
									<td>${error.method}</td>
									<td>${error.url}</td>
									<td>${error.statusCode}</td>
									<td>${error.statusText}</td>
									<td>${error.message}</td>
								</tr>
							`,
								)
								.join('')}
						</tbody>
					</table>
				`,
				)
				.join('')}
		</body>
		</html>
	`
}

const getMailTemplate = (
	capturedErrors: CapturedErrorsType,
	totalErrors: number,
	companyHumanizedName: string,
) => {
	const subject = getMailSubjectTemplate(
		totalErrors,
		companyHumanizedName,
	)
	const html = getMailHtmlTemplate(
		capturedErrors,
		totalErrors,
	)
	return { subject, html }
}

Cypress.on('uncaught:exception', (err, _) => {
	// Prevent Cypress from failing the test on HTTP 500 errors
	if (err.message.includes('500')) {
		// Returning false prevents Cypress from failing the test
		return false
	}
	return true
})

describe('smoke tests', () => {
	beforeEach(() => {
		Cypress.config('defaultCommandTimeout', DEFAULT_TIMEOUT)
		Cypress.config('responseTimeout', DEFAULT_TIMEOUT)
	})

	Object.entries(logins).forEach(([companyName, login]) => {
		const companyNameEnv = Cypress.env('companyName')
		if (companyNameEnv && companyNameEnv !== companyName) {
			return
		}

		it(`test ${companyName}`, () => {
			resetTerminalLogs()
			cy.task('log', `start - ${Cypress.currentTest.title}`)

			goToHomePage(companyName, login)

			const homeRegex = new RegExp(`.*/${companyName}/?$`)
			cy
				.get('.sidebar a')
				.each(($link) => {
					cy
						.wrap($link)
						.invoke('attr', 'href')
						.then((href) => {
							if (href) {
								const shouldSkipLink = (
									href,
									homeRegex,
									includePages,
									excludePages,
								) => {
									// We don't need to check the home again
									const isHomePage = homeRegex.test(href)
									const isNotIncluded =
										includePages.length > 0 &&
										!includePages.some((page) => href.includes(page))
									const isExcluded =
										excludePages.length > 0 &&
										excludePages.some((page) => href.includes(page))

									return isHomePage || isNotIncluded || isExcluded
								}

								if (
									shouldSkipLink(
										href,
										homeRegex,
										SIDEBAR_PAGES_TO_INCLUDE,
										SIDEBAR_PAGES_TO_EXCLUDE,
									)
								) {
									return
								}

								waitForNetworkIdlePrepare()

								addTerminalLog(
									`execution - test ${companyName} - clicked on: ${href}`,
								)

								cy.wrap($link).click({ force: true })
								Object.entries(
									SIDEBAR_PAGES_CUSTOM_ACTIONS,
								).forEach(([page, action]) => {
									if (href.includes(page)) {
										action()
									}
								})

								waitForNetworkIdle()
							}
						})
				})
				.then(() => {
					const testTitle = Cypress.currentTest.title
					const capturedErrors = getCapturedErrors()

					addTerminalLog(`debug - testTitle: ${testTitle}`)
					addTerminalLog(
						`debug - capturedErrors: ${JSON.stringify(capturedErrors)}`,
					)

					if (testTitle in capturedErrors) {
						expect(capturedErrors[testTitle].length).to.be.eq(0)
					}
				})
		})
	})

	afterEach(() => {
		const testTitle = Cypress.currentTest.title
		addTerminalLog(`end - ${testTitle}`)
		logAllTerminalLogs()
	})

	after(() => {
		const capturedErrors = getCapturedErrors()
		const totalErrors = getTotalCapturedErrors(capturedErrors)

		if (totalErrors > 0) {
			const companyName = Cypress.env('companyName')
			const companyHumanizedName =
				logins[companyName]?.companyHumanizedName || companyName
			const mailTemplate = getMailTemplate(
				capturedErrors,
				totalErrors,
				companyHumanizedName,
			)
			cy.task('sendEmail', mailTemplate)
		}
	})
})
