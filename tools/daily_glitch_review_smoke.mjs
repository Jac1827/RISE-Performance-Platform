#!/usr/bin/env node

import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const browserClientPath = "/Users/jacheflin/.codex/plugins/cache/openai-bundled/browser-use/0.1.0-alpha1/scripts/browser-client.mjs";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const pageUrls = {
  index: pathToFileURL(path.join(repoRoot, "docs/portfolio-operations-dashboard/index.html")).href,
  financial: pathToFileURL(path.join(repoRoot, "docs/portfolio-operations-dashboard/financial-accountability.html")).href,
};

function fail(message) {
  throw new Error(message);
}

async function requireVisible(locator, label) {
  if (!(await locator.isVisible())) {
    fail(`${label} was not visible`);
  }
}

async function requireCountAtLeast(locator, label, minimum) {
  const count = await locator.count();
  if (count < minimum) {
    fail(`${label} count was ${count}, expected at least ${minimum}`);
  }
  return count;
}

async function requireNoConsoleErrors(tab, label) {
  const errors = await tab.dev.logs({ levels: ["error"], limit: 50 });
  if (errors.length > 0) {
    const lines = errors.map((entry) => `[${entry.level}] ${entry.message}`).join("\n");
    fail(`${label} emitted console errors:\n${lines}`);
  }
}

async function verifyIndexPage(tab) {
  await tab.goto(pageUrls.index);
  await tab.playwright.waitForLoadState({ state: "load", timeoutMs: 15000 });

  await requireNoConsoleErrors(tab, "index.html");
  await requireVisible(tab.playwright.locator('img[src*="atlas-rise-logo"]'), "ATLAS logo");
  await requireVisible(tab.playwright.locator('img[src*="atlas-rise-tagline"]'), "tagline image");
  await requireVisible(tab.playwright.getByRole("button", { name: "Save Data" }), "Save Data button");
  await requireVisible(tab.playwright.getByRole("button", { name: "Portfolio Home" }), "Portfolio Home tab");
  await requireVisible(tab.playwright.getByRole("button", { name: "Community Setup" }), "Community Setup tab");
  await requireVisible(tab.playwright.getByRole("button", { name: "Data Import" }), "Data Import tab");
  await requireVisible(tab.playwright.getByText("Portfolio Home", { exact: false }), "Portfolio Home text");
  await requireVisible(tab.playwright.getByText("Community Setup", { exact: false }), "Community Setup text");
  await requireVisible(tab.playwright.getByText("Data Import", { exact: false }), "Data Import text");
  await requireCountAtLeast(tab.playwright.locator("button.tab"), "index workspace tabs", 8);
}

async function verifyFinancialPage(tab) {
  await tab.goto(pageUrls.financial);
  await tab.playwright.waitForLoadState({ state: "load", timeoutMs: 15000 });

  await requireNoConsoleErrors(tab, "financial-accountability.html");
  await requireVisible(tab.playwright.locator('img[src*="rise-wordmark-blue"]'), "financial page logo");
  await requireVisible(tab.playwright.getByRole("button", { name: "Import Financial Files" }), "Import Financial Files button");
  await requireVisible(tab.playwright.getByRole("button", { name: "Export Scorecards CSV" }), "Export Scorecards CSV button");
  await requireVisible(tab.playwright.getByRole("button", { name: "Export Board Summary" }), "Export Board Summary button");
  await requireVisible(tab.playwright.getByRole("button", { name: "Export Ownership Report" }), "Export Ownership Report button");
  await requireVisible(tab.playwright.getByRole("button", { name: "Export Board Deck" }), "Export Board Deck button");
  await requireVisible(tab.playwright.getByText("Financial accountability", { exact: false }), "Financial accountability tab text");
  await requireVisible(tab.playwright.getByText("Operational drivers", { exact: false }), "Operational drivers tab text");
  await requireVisible(tab.playwright.getByText("Leasing drivers", { exact: false }), "Leasing drivers tab text");
  await requireCountAtLeast(tab.playwright.locator("button.tab-btn"), "financial layer tabs", 3);
}

async function main() {
  const { setupAtlasRuntime } = await import(browserClientPath);
  await setupAtlasRuntime({ globals: globalThis, backend: "iab" });
  await agent.browser.nameSession("🔎 daily glitch review smoke");

  const indexTab = await agent.browser.tabs.new();
  await verifyIndexPage(indexTab);

  const financialTab = await agent.browser.tabs.new();
  await verifyFinancialPage(financialTab);

  console.log("Daily Glitch Review smoke passed for index.html and financial-accountability.html.");
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
