// Capture the actual deployment, in a fresh browser context with sample data.
import { createRequire } from 'node:module';
import { mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
const require = createRequire(import.meta.url);
const { chromium } = require('../web/node_modules/playwright');
const output = resolve('output/enterprise-manual/screenshots');
mkdirSync(output, { recursive: true });
const baseURL = process.env.E2E_BASE_URL || 'http://192.168.1.11:18860';
const browser = await chromium.launch({ channel: 'chrome' });
try {
  const context = await browser.newContext({ baseURL, viewport: { width: 1280, height: 900 }, deviceScaleFactor: 1 });
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  const shot = async name => {
    await page.locator('body').ariaSnapshot();
    await page.screenshot({ path: resolve(output, `${name}.png`), animations: 'disabled', fullPage: true });
  };
  await page.goto('/');
  await page.getByLabel('账号', {exact:true}).waitFor();
  await shot('login');
  await page.getByLabel('账号', {exact:true}).fill('operator');
  await page.getByLabel('密码', {exact:true}).fill('local-input-only');
  await page.getByRole('button', { name: '登录', exact: true }).click();
  await page.getByRole('heading', { name: '运维总览', exact: true }).waitFor();
  await shot('overview');
  for (const [name, path] of [['assets', '/inventory'], ['metrics', '/observations'], ['incident', '/risk/INC-1201'], ['assistant', '/intelligence'], ['automation', '/automation'], ['executions', '/executions'], ['knowledge', '/library/kb-1'], ['about', '/about']]) {
    await page.goto(path);
    await page.locator('.ex-content h1').waitFor();
    if (name === 'assets') await page.getByRole('searchbox', { name: '搜索资源' }).fill('order');
    if (name === 'assistant') await page.getByRole('button', { name: '高 CPU 应急排查', exact: true }).click();
    await shot(name);
  }
  for (const category of ['overview','hosts','applications','databases','network','clusters','availability','logs','rules']) {
    await page.goto(`/monitoring/${category}`);
    await page.locator('.ex-content h1').waitFor();
    if (category === 'logs') await page.getByRole('searchbox', {name:'搜索日志'}).fill('上游请求超时');
    else if (category !== 'overview' && category !== 'rules') await page.getByRole('searchbox', {name:'搜索监控对象'}).fill('finance');
    await shot(`monitor-${category}`);
  }
  await page.goto('/overview');
  await page.getByRole('button', { name: '账户菜单' }).click();
  await page.getByRole('menuitem', { name: '恢复示例数据', exact: true }).waitFor();
  await shot('account');
  if (errors.length) throw new Error(errors.join('\n'));
  console.log(`Captured ${output}`);
} finally { await browser.close(); }
