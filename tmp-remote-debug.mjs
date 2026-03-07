import { chromium } from '@playwright/test';

const url = process.argv[2] || 'https://kake-git-hub.github.io/LessonScheduleTable/#/c/9998/admin/2026-spring';
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage();
page.on('console', (msg) => console.log('[console]', msg.type(), msg.text()));
page.on('pageerror', (err) => console.log('[pageerror]', err?.stack || String(err)));
page.on('requestfailed', (req) => console.log('[requestfailed]', req.url(), req.failure()?.errorText || 'unknown'));
await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 120000 });
await page.waitForTimeout(12000);
console.log('[url]', page.url());
console.log('[body]', (await page.textContent('body'))?.slice(0, 2000));
await browser.close();
