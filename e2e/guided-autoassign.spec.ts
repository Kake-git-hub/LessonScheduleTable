import { expect, test, type Page } from '@playwright/test'
import { createClassroom, deleteClassroom, futureDate, waitForAuth } from './helpers/auth'

const CLASSROOM_ID = `e2eguide${Date.now()}`
const CLASSROOM_NAME = 'E2E誘導テスト教室'
const ADMIN_PASSWORD = 'admin1234'
const START_DATE = futureDate(7)
const END_DATE = futureDate(7)

test.describe.skip('Guided auto-assign flow', () => {
	test.afterAll(async ({ browser }) => {
		const page = await browser.newPage()
		await deleteClassroom(page, CLASSROOM_NAME).catch(() => {})
		await page.close()
	})

	test('講師不足があると自動提案は解消導線へ切り替わる', async ({ page }) => {
		page.on('dialog', (dialog) => void dialog.accept())

		await createClassroom(page, CLASSROOM_ID, CLASSROOM_NAME)
		await openClassroom(page)
		await addTeacher(page, '誘導先生')
		await addStudent(page, '誘導生徒', '中3')
		await createLectureSession(page)
		await openAdminPage(page)

		await openDirectInput(page, '誘導先生')
		await toggleFirstTeacherSlot(page)
		await submitInput(page)

		await openAdminPage(page)
		await openDirectInput(page, '誘導生徒')
		await addStudentSubjectRequest(page, '数', '1')
		await submitInput(page)

		await openAdminPage(page)
		await page.getByRole('button', { name: '自動提案' }).click()
		await expect(page.getByRole('button', { name: '結果詳細' })).toBeVisible({ timeout: 15000 })

		await openAdminPage(page)
		await openDirectInput(page, '誘導先生')
		await clearSelectedTeacherSlots(page)
		await submitInput(page)

		await openAdminPage(page)
		await expect(page.getByRole('button', { name: '未解決を先に解消' })).toBeVisible({ timeout: 15000 })
		await page.getByRole('button', { name: '未解決を先に解消' }).click()
		await expect(page.locator('text=講師不足')).toBeVisible({ timeout: 15000 })
	})
})

async function openClassroom(page: Page): Promise<void> {
	await page.goto(`/#/c/${CLASSROOM_ID}`)
	await waitForAuth(page)
	const continueButton = page.getByRole('button', { name: '続行' })
	if (await continueButton.isVisible({ timeout: 3000 }).catch(() => false)) {
		await page.locator('input[placeholder="管理者パスワード"]').fill(ADMIN_PASSWORD)
		await continueButton.click()
	}
	await expect(page.locator('text=講習コマ割りアプリ')).toBeVisible({ timeout: 15000 })
}

async function addTeacher(page: Page, teacherName: string): Promise<void> {
	await openClassroom(page)
	await page.locator('input[placeholder="講師名"]').fill(teacherName)
	await page.getByRole('checkbox').first().check()
	const teacherPanel = page.locator('.panel', { has: page.locator('h3:has-text("講師登録")') })
	await teacherPanel.getByRole('button', { name: '追加' }).click()
	await expect(page.locator('td', { hasText: teacherName })).toBeVisible({ timeout: 15000 })
}

async function addStudent(page: Page, studentName: string, grade: string): Promise<void> {
	await openClassroom(page)
	await page.locator('input[placeholder="生徒名"]').fill(studentName)
	const studentPanel = page.locator('.panel', { has: page.locator('h3:has-text("生徒登録")') })
	await studentPanel.locator('select').first().selectOption(grade)
	await studentPanel.getByRole('button', { name: '追加' }).click()
	await expect(page.locator('td', { hasText: studentName })).toBeVisible({ timeout: 15000 })
}

async function createLectureSession(page: Page): Promise<void> {
	await openClassroom(page)
	const sessionPanel = page.locator('.panel', { has: page.locator('h3:has-text("新規特別講習を追加")') })
	const dateInputs = sessionPanel.locator('input[type="date"]')
	await dateInputs.nth(0).fill(START_DATE)
	await dateInputs.nth(1).fill(END_DATE)
	await sessionPanel.getByRole('button', { name: '特別講習を作成' }).click()
	await expect(page.locator('tr', { has: page.locator('td', { hasText: 'summer' }) })).toBeVisible({ timeout: 15000 })
}

async function openAdminPage(page: Page): Promise<void> {
	await openClassroom(page)
	const sessionRow = page.locator('tr', { has: page.locator('td', { hasText: 'summer' }) }).first()
	await sessionRow.getByRole('button', { name: '管理' }).click()
	await waitForAuth(page)
	await expect(page.getByRole('button', { name: /自動提案|未解決を先に解消/ })).toBeVisible({ timeout: 15000 })
}

async function openDirectInput(page: Page, personName: string): Promise<void> {
	const row = page.locator('tr', { has: page.locator('td', { hasText: personName }) }).first()
	await row.getByRole('link', { name: /直接入力/ }).click()
	await waitForAuth(page)
}

async function toggleFirstTeacherSlot(page: Page): Promise<void> {
	const slotButton = page.locator('button.teacher-slot-btn:not([disabled])').first()
	await slotButton.click()
}

async function clearSelectedTeacherSlots(page: Page): Promise<void> {
	const activeButtons = page.locator('button.teacher-slot-btn.active')
	const count = await activeButtons.count()
	for (let index = 0; index < count; index += 1) {
		await activeButtons.nth(0).click()
	}
}

async function addStudentSubjectRequest(page: Page, subject: string, count: string): Promise<void> {
	const selects = page.locator('select')
	await selects.filter({ has: page.locator('option', { hasText: '＋ 科目を追加' }) }).first().selectOption(subject)
	await page.locator('input[type="number"]').first().fill(count)
}

async function submitInput(page: Page): Promise<void> {
	const downloadPromise = page.waitForEvent('download', { timeout: 10000 }).catch(() => null)
	await page.getByRole('button', { name: '送信' }).click()
	await downloadPromise
	await expect(page.locator('text=入力完了')).toBeVisible({ timeout: 15000 })
}
