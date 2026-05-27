const { test, expect } = require('@playwright/test');

test('backend returns 404 for unknown job', async ({ request }) => {
  const resp = await request.get('http://localhost:4000/results/999999');
  expect(resp.status()).toBe(404);
});
