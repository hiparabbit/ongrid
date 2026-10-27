import { test, expect, type Page } from '@playwright/test';
async function enter(page: Page) {
  await page.goto('/');
  await page.getByRole('button', { name: '进入演示', exact: true }).click();
  await expect(
    page.getByRole('heading', { name: '运维总览', exact: true }),
  ).toBeVisible();
}
async function account(page: Page, action: string) {
  await page.getByRole('button', { name: '账户菜单' }).click();
  await page.getByRole('menuitem', { name: action, exact: true }).click();
}
test('complete response, editing, cancellation, persistence and reset without API calls', async ({
  page,
}) => {
  const api: string[] = [];
  page.on('request', (request) => {
    if (/^\/(api|v1)(\/|$)/.test(new URL(request.url()).pathname))
      api.push(request.url());
  });
  await enter(page);
  await page.goto('/risk/INC-1201');
  await page.getByRole('button', { name: '确认告警', exact: true }).click();
  await expect(
    page.getByRole('button', { name: '确认告警', exact: true }),
  ).toBeDisabled();
  await page.reload();
  await expect(
    page.getByRole('button', { name: '确认告警', exact: true }),
  ).toBeDisabled();
  await page.getByRole('button', { name: '关闭告警', exact: true }).click();
  await page.getByRole('button', { name: '取消', exact: true }).click();
  await expect(
    page.getByRole('button', { name: '关闭告警', exact: true }),
  ).toBeEnabled();
  await page.getByRole('link', { name: 'AI 辅助分析', exact: true }).click();
  await page
    .getByRole('button', { name: '高 CPU 应急排查', exact: true })
    .click();
  await expect(
    page.getByText('资源 order-api-01 的 CPU 快照为 94%', { exact: false }),
  ).toBeVisible();
  await page.getByRole('link', { name: '前往关联流程', exact: true }).click();
  const flow = page.locator('.ex-flow').filter({
    has: page.getByRole('heading', { name: '高 CPU 应急排查', exact: true }),
  });
  await flow.getByRole('button', { name: '模拟执行', exact: true }).click();
  await page.getByRole('button', { name: '取消', exact: true }).click();
  await page.goto('/executions');
  await expect(page.locator('tbody tr')).toHaveCount(6);
  await page.goto('/automation?flow=cpu');
  await flow.getByRole('button', { name: '模拟执行', exact: true }).click();
  await page.getByRole('button', { name: '确认', exact: true }).click();
  await page.getByRole('link', { name: /^查看记录/ }).click();
  await expect(page.locator('tbody tr')).toHaveCount(7);
  await page.goto('/library/kb-1');
  await page.getByRole('button', { name: '编辑文档', exact: true }).click();
  await page.getByLabel('文档标题', { exact: true }).fill('操作手册验证指南');
  await page.getByRole('button', { name: '保存文档', exact: true }).click();
  await page.reload();
  await expect(
    page.getByRole('heading', { name: '操作手册验证指南', exact: true }),
  ).toBeVisible();
  await account(page, '恢复示例数据');
  await page.getByRole('button', { name: '取消', exact: true }).click();
  await expect(
    page.getByRole('heading', { name: '操作手册验证指南', exact: true }),
  ).toBeVisible();
  await account(page, '恢复示例数据');
  await page.getByRole('button', { name: '确认', exact: true }).click();
  await expect(
    page.getByRole('heading', { name: '高 CPU 排查指南', exact: true }),
  ).toBeVisible();
  await page.goto('/settings/llm');
  await expect(page).toHaveURL(/\/overview$/);
  expect(api).toEqual([]);
});

for (const width of [390, 768, 1280, 1920])
  for (const lang of ['zh-CN', 'en-US'])
    for (const theme of ['light', 'dark']) {
      test(`layout ${width} ${lang} ${theme}`, async ({ page }) => {
        await page.setViewportSize({ width, height: 1000 });
        await page.addInitScript(
          ({ lang, theme }) => {
            localStorage.setItem('ongrid-locale', lang);
            localStorage.setItem('enterprise-ai.theme', theme);
            sessionStorage.setItem('enterprise-ai.entered', '1');
          },
          { lang, theme },
        );
        for (const path of [
          '/overview',
          '/risk',
          '/risk/INC-1201',
          '/inventory',
          '/inventory/res-1',
          '/observations',
          '/intelligence',
          '/automation',
          '/executions',
          '/library',
          '/library/kb-1',
          '/about',
        ]) {
          await page.goto(path);
          await expect(page.locator('.ex-content h1')).toBeVisible();
          // A long resource name exercises wrapping and table boundaries.
          if (path === '/inventory') {
            await page.evaluate(() => {
              const data = JSON.parse(
                localStorage.getItem('enterprise-ai.sandbox.v1') || 'null',
              );
              if (data) {
                data.resources[0].name =
                  'enterprise-production-resource-very-long-name-'.repeat(6);
                localStorage.setItem(
                  'enterprise-ai.sandbox.v1',
                  JSON.stringify(data),
                );
              }
            });
            await page.reload();
          }
          await page.locator('.ex-footer').scrollIntoViewIfNeeded();
          await expect(page.locator('.ex-footer')).toBeInViewport();
          const bounds = await page.evaluate(() => ({
            viewport: innerWidth,
            document: document.documentElement.scrollWidth,
            content: document
              .querySelector('.ex-content')!
              .getBoundingClientRect().right,
          }));
          expect(bounds.document, path).toBeLessThanOrEqual(
            bounds.viewport + 1,
          );
          expect(bounds.content, path).toBeLessThanOrEqual(bounds.viewport + 1);
          for (const table of await page.locator('.ex-table-scroll').all()) {
            await table.evaluate((el) => {
              el.scrollLeft = el.scrollWidth;
            });
            const last = table.locator('tbody tr').first().locator('td').last();
            const cellBounds = await last.boundingBox();
            const tableBounds = await table.boundingBox();
            expect(cellBounds!.x + cellBounds!.width).toBeLessThanOrEqual(
              tableBounds!.x + tableBounds!.width + 2,
            );
          }
        }
        await page.goto('/overview');
        await page.screenshot({
          path: `../output/playwright/enterprise-${width}-${lang}-${theme}.png`,
          fullPage: true,
        });
        if (width <= 768) {
          await page
            .getByRole('button', {
              name: lang === 'zh-CN' ? '打开导航' : 'Open navigation',
            })
            .click();
          await expect(page.getByRole('dialog')).toBeVisible();
          await page.keyboard.press('Escape');
          await expect(page.getByRole('dialog')).not.toBeVisible();
        }
      });
    }
