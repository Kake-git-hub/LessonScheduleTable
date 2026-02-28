import type { Page } from '@playwright/test'

/**
 * Wait for Firebase anonymous auth to complete.
 * The app shows "接続中..." while authenticating.
 */
export async function waitForAuth(page: Page): Promise<void> {
  await page.waitForFunction(
    () => !document.body.textContent?.includes('接続中...'),
    { timeout: 15000 },
  ).catch(() => {
    // If timeout, continue — auth may have completed differently
  })
}

/**
 * Create a classroom via the UI on the ClassroomSelectPage.
 */
export async function createClassroom(page: Page, id: string, name: string): Promise<void> {
  await page.goto('/#/')
  await waitForAuth(page)

  // Clean up any leftover classrooms with the same name from previous runs
  page.on('dialog', (dialog) => void dialog.accept())
  while (true) {
    const staleRow = page.locator('tr', { has: page.locator('td strong', { hasText: name }) }).first()
    if (await staleRow.isVisible({ timeout: 1000 }).catch(() => false)) {
      await staleRow.getByRole('button', { name: '削除' }).click()
      await page.waitForTimeout(1000)
    } else {
      break
    }
  }

  await page.locator('input[placeholder="教室ID（英数字）"]').fill(id)
  await page.locator('input[placeholder="教室名"]').fill(name)
  await page.getByRole('button', { name: '作成' }).click()
  // Wait for the classroom to appear in the table (use row with matching ID for precision)
  await page.locator('tr', { has: page.locator('td', { hasText: id }) })
    .locator('td strong', { hasText: name })
    .waitFor({ timeout: 10000 })
}

/**
 * Delete a classroom via the UI (handles confirm dialog).
 */
export async function deleteClassroom(page: Page, name: string): Promise<void> {
  await page.goto('/#/')
  await waitForAuth(page)
  // Find the row with this classroom name and click delete
  const row = page.locator('tr', { has: page.locator('td strong', { hasText: name }) })
  page.on('dialog', (dialog) => void dialog.accept())
  await row.getByRole('button', { name: '削除' }).click()
  // Wait for the classroom to disappear
  await page.locator('td strong', { hasText: name }).waitFor({ state: 'hidden', timeout: 10000 }).catch(() => {})
}

/**
 * Generate a date string offset from today.
 */
export function futureDate(daysFromNow: number): string {
  const d = new Date()
  d.setDate(d.getDate() + daysFromNow)
  return d.toISOString().split('T')[0]
}
