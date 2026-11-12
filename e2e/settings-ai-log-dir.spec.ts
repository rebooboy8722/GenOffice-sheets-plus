import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test('the AI log-folder setting lives in the AI Model pane and persists to ai-settings.json', async () => {
  // per-machine path, set through the UI (the whole point: never hardcoded)
  const logDir = '/tmp/genoffice-ai-log-e2e'
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-ai-log-dir',
  })
  const { page, userDataDir } = launched
  try {
    await page.getByRole('button', { name: 'Settings', exact: true }).click()
    await page.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()

    const input = page.locator('#set-ai-log-dir')
    await expect(input).toBeVisible()
    await expect(input).toHaveAttribute('placeholder', 'Blank = no logging')
    // the native folder picker cannot be driven headless; the text field is
    // the operable path and the button proves the picker wiring exists
    await expect(page.getByRole('button', { name: 'Choose…', exact: true })).toBeVisible()

    await input.fill(logDir)
    // blur commits the draft (trimmed; blank clears the setting)
    await input.blur()
    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(page.locator('.set-pane-actions .set-ai-status.ok')).toHaveText('Saved')
    // the field round-trips from the persisted settings, not from the draft
    await expect(input).toHaveValue(logDir)
    await expect
      .poll(async () => {
        const saved = JSON.parse(await readFile(join(userDataDir, 'ai-settings.json'), 'utf8'))
        return saved.logDir
      })
      .toEqual(logDir)
    await page.screenshot({ path: screenshotPath('settings-ai-log-dir') })
  } finally {
    await closeAndSaveVideo(launched, 'settings-ai-log-dir')
  }
})
