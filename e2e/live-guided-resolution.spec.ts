import { expect, test, type Page } from '@playwright/test'
import { appendFileSync, mkdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { initializeApp } from 'firebase/app'
import { getAuth, signInAnonymously } from 'firebase/auth'
import { deleteDoc, doc, getDoc, getFirestore, setDoc } from 'firebase/firestore'
import { waitForAuth } from './helpers/auth'

const firebaseConfig = {
  apiKey: 'AIzaSyDPs5KUD9j-Oa3lwxEo4so_4tasLlSYI7Q',
  authDomain: 'lessonscheduletable.firebaseapp.com',
  projectId: 'lessonscheduletable',
  storageBucket: 'lessonscheduletable.firebasestorage.app',
  messagingSenderId: '390016573631',
  appId: '1:390016573631:web:82b7d38d9fb53a9938e60a',
}

const SOURCE_CLASSROOM_ID = '9998'
const SOURCE_SESSION_ID = '2026-spring'
const ADMIN_PASSWORD = 'admin1234'

type StatusSectionSummary = {
  title: string
  items: number
  actionCount: number
  texts: string[]
}

const app = initializeApp(firebaseConfig)
const auth = getAuth(app)
const db = getFirestore(app)

test.describe('Live guided resolution flow', () => {
  test('resolved teacher shortages do not reappear after auto-assign', async ({ page }) => {
    test.setTimeout(180000)
    const clonedClassroomId = `e2elive${Date.now()}`
    const logPath = resolve(process.cwd(), 'test-results', `live-guided-resolution-${clonedClassroomId}.log`)
    const log = (label: string, payload?: unknown): void => {
      mkdirSync(dirname(logPath), { recursive: true })
      appendFileSync(logPath, `${label}${payload === undefined ? '' : ` ${JSON.stringify(payload)}`}\n`)
    }

    page.on('dialog', (dialog) => void dialog.accept())

    await cloneSession(clonedClassroomId)

    try {
      log('openAdmin:start', { clonedClassroomId })
      await openAdminPage(page, clonedClassroomId, SOURCE_SESSION_ID)

      for (let step = 0; step < 10; step += 1) {
        const resolveVisible = await page.getByRole('button', { name: /未解決: \d+件/ }).isVisible().catch(() => false)
        log('blockerResolveVisible', { step, resolveVisible })
        if (!resolveVisible) break
        const modalOpened = await ensureStatusModalOpen(page)
        if (!modalOpened) {
          log('blockerAutoAssignTriggered', { step })
          break
        }
        const sections = await readStatusSections(page)
        log(`blockerStep:${step}`, sections)
        const actionResult = await clickFirstAction(page, ['講師不足', '過割当'])
        log('blockerActed', { step, ...actionResult })
        expect(actionResult.acted).toBeTruthy()
        await page.waitForTimeout(1200)

        const modalVisible = await page.locator('.status-modal-backdrop').isVisible().catch(() => false)
        if (!modalVisible) {
          log(`blockerStepClosed:${step}`)
          continue
        }

        const refreshedSections = await readStatusSections(page)
        log(`blockerStepRefreshed:${step}`, refreshedSections)
        const shortageCount = getSectionItemCount(refreshedSections, '講師不足')
        expect(shortageCount).toBeLessThanOrEqual(getSectionItemCount(sections, '講師不足'))
      }

      await closeStatusModal(page)
      log('autoAssign:click')
      const proposalModalOpened = await openStatusFlow(page)

      if (proposalModalOpened) {
        for (let step = 0; step < 20; step += 1) {
          const reopened = await openStatusFlow(page)
          expect(reopened).toBeTruthy()
          const sections = await readStatusSections(page)
          log(`proposalStep:${step}`, sections)
          expect(getSectionItemCount(sections, '講師不足')).toBe(0)

          const actionResult = await clickFirstAction(page, ['過割当', '残コマ詳細', '残コマ', '過割当解除'])
          log('proposalActed', { step, ...actionResult })
          if (!actionResult.acted) {
            break
          }
          await page.waitForTimeout(1200)
        }

        const finalModalOpened = await openStatusFlow(page)
        if (finalModalOpened) {
          const finalSections = await readStatusSections(page)
          log('finalSections', finalSections)
          expect(getSectionItemCount(finalSections, '講師不足')).toBe(0)
          expect(finalSections.some((section) => section.title === '講師不足')).toBeFalsy()
        }
      }
    } finally {
      log('cleanup:start')
      await cleanupSession(clonedClassroomId)
      log('cleanup:done')
    }
  })
})

async function cloneSession(clonedClassroomId: string): Promise<void> {
  await signInAnonymously(auth)
  const sourceRef = doc(db, 'classrooms', SOURCE_CLASSROOM_ID, 'sessions', SOURCE_SESSION_ID)
  const sourceSnap = await getDoc(sourceRef)
  if (!sourceSnap.exists()) throw new Error('source session not found')
  const sourceMasterRef = doc(db, 'classrooms', SOURCE_CLASSROOM_ID, 'master', 'default')
  const sourceMasterSnap = await getDoc(sourceMasterRef)

  await setDoc(doc(db, 'classrooms', clonedClassroomId), {
    name: `E2E Live ${clonedClassroomId}`,
    createdAt: Date.now(),
  })
  if (sourceMasterSnap.exists()) {
    await setDoc(doc(db, 'classrooms', clonedClassroomId, 'master', 'default'), sourceMasterSnap.data())
  }
  await setDoc(doc(db, 'classrooms', clonedClassroomId, 'sessions', SOURCE_SESSION_ID), sourceSnap.data())
}

async function cleanupSession(clonedClassroomId: string): Promise<void> {
  await deleteDoc(doc(db, 'classrooms', clonedClassroomId, 'sessions', SOURCE_SESSION_ID)).catch(() => {})
  await deleteDoc(doc(db, 'classrooms', clonedClassroomId, 'master', 'default')).catch(() => {})
  await deleteDoc(doc(db, 'classrooms', clonedClassroomId)).catch(() => {})
}

async function openAdminPage(page: Page, classroomId: string, sessionId: string): Promise<void> {
  await page.goto(`/#/c/${classroomId}/admin/${sessionId}`)
  await waitForAuth(page)

  const passwordInput = page.locator('input[placeholder="管理者パスワード"]')
  if (await passwordInput.isVisible({ timeout: 3000 }).catch(() => false)) {
    await passwordInput.fill(ADMIN_PASSWORD)
    await page.getByRole('button', { name: '続行' }).click()
  }

  await expect(page.locator('body')).not.toContainText('空の特別講習を作成', { timeout: 10000 })
  await expect.poll(async () => {
    if (await page.getByRole('button', { name: '自動提案' }).isVisible().catch(() => false)) return 'auto'
    if (await page.getByRole('button', { name: /未解決: \d+件/ }).isVisible().catch(() => false)) return 'resolve'
    return ''
  }, { timeout: 30000 }).not.toBe('')
}

async function openStatusFlow(page: Page): Promise<boolean> {
  if (await ensureStatusModalOpen(page)) return true
  await page.waitForTimeout(1200)
  return ensureStatusModalOpen(page)
}

async function ensureStatusModalOpen(page: Page): Promise<boolean> {
  const modal = page.locator('.status-modal-backdrop')
  if (await modal.isVisible().catch(() => false)) return true

  const resolveButton = page.getByRole('button', { name: /未解決: \d+件/ })
  const autoButton = page.getByRole('button', { name: '自動提案' })

  await expect.poll(async () => {
    if (await modal.isVisible().catch(() => false)) return 'modal'
    if (await resolveButton.isVisible().catch(() => false)) return 'resolve'
    if (await autoButton.isVisible().catch(() => false)) return 'auto'
    return ''
  }, { timeout: 20000 }).not.toBe('')

  if (await modal.isVisible().catch(() => false)) return true

  if (await resolveButton.isVisible().catch(() => false)) {
    await resolveButton.click()
    await page.waitForTimeout(1200)
    return modal.isVisible().catch(() => false)
  }

  if (await autoButton.isVisible().catch(() => false)) {
    await autoButton.click()
    await page.waitForTimeout(1200)
    return modal.isVisible().catch(() => false)
  }

  throw new Error('status modal controls did not become available')
}

async function closeStatusModal(page: Page): Promise<void> {
  const modal = page.locator('.status-modal-backdrop')
  if (!await modal.isVisible().catch(() => false)) return
  await page.getByRole('button', { name: '閉じる', exact: true }).click()
  await expect(modal).toBeHidden({ timeout: 15000 })
}

async function readStatusSections(page: Page): Promise<StatusSectionSummary[]> {
  await expect(page.locator('.status-modal-backdrop')).toBeVisible({ timeout: 15000 })
  return page.evaluate(() => {
    return [...document.querySelectorAll('.status-section')].map((section) => ({
      title: section.querySelector('h4')?.textContent?.trim() ?? '',
      items: section.querySelectorAll('.status-item-card').length,
      actionCount: section.querySelectorAll('.status-proposal-action').length,
      texts: [...section.querySelectorAll('.status-item-card')].map((item) => item.textContent?.replace(/\s+/g, ' ').trim() ?? ''),
    }))
  })
}

function getSectionItemCount(sections: StatusSectionSummary[], title: string): number {
  return sections.find((section) => section.title === title)?.items ?? 0
}

async function clickFirstAction(page: Page, titles: string[]): Promise<{ acted: boolean; sectionTitle?: string; itemTitle?: string; buttonText?: string; selectedChoice?: string }> {
  for (const title of titles) {
    const section = page.locator('.status-section').filter({ has: page.locator('h4', { hasText: title }) }).first()
    const actionCount = await section.locator('.status-proposal-action').count().catch(() => 0)
    if (actionCount > 0) {
      const button = section.locator('.status-proposal-action').first()
      const metadata = await section.evaluate((element) => {
        const card = element.querySelector('.status-item-card')
        const actionButton = card?.querySelector<HTMLButtonElement>('.status-proposal-action')
        const select = actionButton?.previousElementSibling instanceof HTMLSelectElement
          ? actionButton.previousElementSibling
          : null
        return {
          itemTitle: card?.querySelector('.status-item-title')?.textContent?.trim() ?? undefined,
          buttonText: actionButton?.textContent?.trim() ?? undefined,
          selectedChoice: select?.selectedOptions?.[0]?.textContent?.trim() ?? undefined,
        }
      }).catch(() => null)
      await button.click()
      return {
        acted: true,
        sectionTitle: title,
        itemTitle: metadata?.itemTitle,
        buttonText: metadata?.buttonText,
        selectedChoice: metadata?.selectedChoice,
      }
    }
  }
  return { acted: false }
}