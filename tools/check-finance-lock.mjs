import { readFileSync } from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const backendPath = path.join(projectRoot, "assents", "motor", "motorbackend.js");

const expected = {
  "2026-01": {
    total: 956548.886,
    nfs: ["721", "722", "725", "726", "727", "728", "729", "731"]
  },
  "2026-02": {
    total: 98509.9724,
    nfs: ["734", "736", "737", "741", "742", "743", "744", "746", "747", "748", "749", "750", "751"]
  },
  "2026-03": {
    total: 567327.6052,
    nfs: ["752", "754", "755", "756", "757", "758", "759", "760", "761", "762", "767", "768", "769", "770", "771", "775"]
  },
  "2026-04": {
    total: 603915.9832,
    nfs: ["776", "777", "780", "781", "785", "788", "789", "790", "791", "792"]
  }
};

function parseDateBR(value) {
  const match = String(value || "").trim().match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return "";
  return `${match[3]}-${match[2]}`;
}

function normalizeNF(value) {
  return String(value || "").replace(/\D/g, "");
}

function assertEqual(label, actual, expectedValue) {
  if (actual !== expectedValue) {
    throw new Error(`${label}: esperado ${expectedValue}, recebido ${actual}`);
  }
}

function assertArray(label, actual, expectedValue) {
  const actualText = actual.join(",");
  const expectedText = expectedValue.join(",");
  assertEqual(label, actualText, expectedText);
}

const source = readFileSync(backendPath, "utf8");
const context = {
  window: {},
  console,
  fetch: async () => {
    throw new Error("check-finance-lock nao deve acessar rede.");
  }
};

vm.runInNewContext(
  `${source}\n;globalThis.__financeLock = { FATURAMENTO_FINANCEIRO_2026, MESES_FINANCEIRO_FECHADOS_2026 };`,
  context,
  { filename: "motorbackend.js" }
);

const finance = context.__financeLock;
const locked = finance.FATURAMENTO_FINANCEIRO_2026;
const closedMonths = Array.from(finance.MESES_FINANCEIRO_FECHADOS_2026).sort();
const expectedMonths = Object.keys(expected).sort();

assertArray("Meses financeiros fechados", closedMonths, expectedMonths);

for (const month of expectedMonths) {
  const docs = Object.entries(locked)
    .map(([nf, doc]) => ({ nf: normalizeNF(nf), doc }))
    .filter(({ doc }) => parseDateBR(doc && doc.data) === month)
    .filter(({ doc }) => !(doc && doc.contabiliza === false));

  const nfs = docs.map(({ nf }) => nf).sort((a, b) => a.localeCompare(b, "pt-BR", { numeric: true }));
  const total = docs.reduce((sum, { doc }) => sum + Number(doc.valor || 0), 0);

  assertArray(`NFs contabilizaveis ${month}`, nfs, expected[month].nfs);

  if (Math.abs(total - expected[month].total) > 0.01) {
    throw new Error(`Total ${month}: esperado ${expected[month].total}, recebido ${total}`);
  }
}

console.log("[finance-lock] Snapshot financeiro fechado validado.");
