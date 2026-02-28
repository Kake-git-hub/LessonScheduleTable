import { test, expect } from '@playwright/test'

test('app loads and shows classroom selector', async ({ page }) => {
  await page.goto('/')
  // Wait for Firebase auth to complete (loading state disappears)
  await page.waitForFunction(() => {
    return !document.body.textContent?.includes('接続中...')
  }, { timeout: 15000 }).catch(() => {
    // If timeout, continue anyway — auth may have completed differently
  })

  // The classroom selection page should be visible
  await expect(page.locator('body')).toContainText('教室', { timeout: 15000 })
})

test('classroom create form is visible', async ({ page }) => {
  await page.goto('/')
  // Wait for page to load
  await page.waitForFunction(() => {
    return !document.body.textContent?.includes('接続中...')
  }, { timeout: 15000 }).catch(() => {})

  // Should have an input or form to create a classroom
  const createButton = page.getByRole('button', { name: /作成|追加|新規/i })
  await expect(createButton).toBeVisible({ timeout: 10000 })
})

test('version badge is displayed', async ({ page }) => {
  await page.goto('/')
  await page.waitForFunction(() => {
    return !document.body.textContent?.includes('接続中...')
  }, { timeout: 15000 }).catch(() => {})

  // App version should be shown somewhere
  await expect(page.locator('body')).toContainText('1.0.0', { timeout: 10000 })
})
