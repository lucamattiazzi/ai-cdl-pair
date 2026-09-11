import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";

const browser = await chromium.launch();
try {
  const svg = await readFile("apps/pair-addin/public/icon.svg", "utf8");
  for (const size of [16, 32, 64, 80, 128]) {
    const page = await browser.newPage({
      viewport: { width: size, height: size },
      deviceScaleFactor: 1,
    });
    await page.setContent(
      `<style>body{margin:0}svg{display:block;width:100%;height:100%}</style>${svg}`,
    );
    await page.screenshot({
      path: `apps/pair-addin/public/icon-${size}.png`,
      omitBackground: true,
    });
    await page.close();
  }
} finally {
  await browser.close();
}
