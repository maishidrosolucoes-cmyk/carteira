import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const backendPath = path.join(projectRoot, "assents", "motor", "motorbackend.js");
const checkPath = path.join(projectRoot, "tools", "check-finance-lock.mjs");

function parseArgs(argv) {
  return argv.reduce((acc, arg) => {
    if (arg.startsWith("--") && arg.includes("=")) {
      const [key, ...rest] = arg.slice(2).split("=");
      acc[key] = rest.join("=");
    } else if (arg.startsWith("--")) {
      acc[arg.slice(2)] = true;
    }
    return acc;
  }, {});
}

function fail(message) {
  console.error(`[finance-lock] ${message}`);
  process.exit(1);
}

function normalizeDigits(value) {
  return String(value || "").replace(/\D/g, "");
}

function parseMoney(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let str = String(value).trim();
  if (!str) return 0;

  str = str.replace(/\s/g, "").replace(/[R$r$\u00A0]/g, "");
  if (str.includes(",")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if ((str.match(/\./g) || []).length > 1) {
    str = str.replace(/\./g, "");
  }

  const parsed = Number.parseFloat(str.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeMonth(value) {
  const raw = String(value || "").trim();
  const match = raw.match(/^(\d{4})-(\d{2})$/);
  if (!match) fail("Informe o mes no formato YYYY-MM. Exemplo: npm run finance:lock -- --month=2026-05");

  const month = Number(match[2]);
  if (!Number.isInteger(month) || month < 1 || month > 12) fail(`Mes invalido: ${raw}`);
  return raw;
}

function getPreviousMonthKey() {
  const now = new Date();
  const year = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const month = now.getMonth() === 0 ? 12 : now.getMonth();
  return `${year}-${String(month).padStart(2, "0")}`;
}

function getMonthKey(value) {
  const raw = String(value || "").trim();
  let match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}`;

  match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) return `${match[3]}-${match[2]}`;

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
  }

  return "";
}

function toDateBR(value) {
  const raw = String(value || "").trim();
  let match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[3]}/${match[2]}/${match[1]}`;

  match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) return raw;

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return "";

  return [
    String(parsed.getDate()).padStart(2, "0"),
    String(parsed.getMonth() + 1).padStart(2, "0"),
    String(parsed.getFullYear())
  ].join("/");
}

function splitList(value) {
  return new Set(String(value || "")
    .split(/[,\s;]+/)
    .map(normalizeDigits)
    .filter(Boolean));
}

function loadFinanceContext() {
  const source = readFileSync(backendPath, "utf8");
  const context = {
    window: {},
    console,
    fetch: async () => {
      throw new Error("lock-finance-month controla a leitura da API separadamente.");
    }
  };

  vm.runInNewContext(
    `${source}\n;globalThis.__financeLock = { FATURAMENTO_FINANCEIRO_2026, MESES_FINANCEIRO_FECHADOS_2026, ERP_CARTEIRA_API_URL };`,
    context,
    { filename: "motorbackend.js" }
  );

  return {
    source,
    url: context.__financeLock.ERP_CARTEIRA_API_URL,
    locked: context.__financeLock.FATURAMENTO_FINANCEIRO_2026,
    closedMonths: new Set(Array.from(context.__financeLock.MESES_FINANCEIRO_FECHADOS_2026))
  };
}

function addRowToGroup(groups, row, targetMonth, includeList, excludeList) {
  const nf = normalizeDigits(row && row.nf);
  if (!nf) return;
  if (includeList && !includeList.has(nf)) return;

  const data = toDateBR(row.data_faturam || row.data_faturamento);
  if (!data || getMonthKey(data) !== targetMonth) return;

  const valor = parseMoney(row.p_total);
  if (valor <= 0) return;

  if (!groups.has(nf)) {
    groups.set(nf, {
      nf,
      data,
      valoresCentavos: new Set(),
      valores: [],
      clientes: new Set(),
      contabiliza: !excludeList.has(nf)
    });
  }

  const group = groups.get(nf);
  const cents = String(Math.round(valor * 10000));
  group.valoresCentavos.add(cents);
  group.valores.push(valor);
  if (row.cliente) group.clientes.add(String(row.cliente).trim());
  if (excludeList.has(nf)) group.contabiliza = false;
}

function buildDocsFromGroups(groups) {
  return Array.from(groups.values())
    .sort((a, b) => a.nf.localeCompare(b.nf, "pt-BR", { numeric: true }))
    .map(group => {
      if (group.valoresCentavos.size > 1) {
        fail(`NF ${group.nf} tem mais de um p_total no mesmo mes. Fechamento bloqueado para evitar soma incorreta.`);
      }

      return {
        nf: group.nf,
        doc: {
          data: group.data,
          valor: group.valores[0],
          cliente: Array.from(group.clientes)[0] || "",
          ...(group.contabiliza ? {} : { contabiliza: false })
        }
      };
    });
}

function formatNumber(value) {
  const rounded = Math.round(Number(value) * 10000) / 10000;
  return Number.isInteger(rounded) ? String(rounded) : String(rounded);
}

function formatFinanceObject(locked) {
  const lines = Object.entries(locked)
    .sort(([a], [b]) => a.localeCompare(b, "pt-BR", { numeric: true }))
    .map(([nf, doc]) => {
      const parts = [
        `data: ${JSON.stringify(doc.data || "")}`,
        `valor: ${formatNumber(doc.valor || 0)}`,
        `cliente: ${JSON.stringify(doc.cliente || "")}`
      ];
      if (doc.contabiliza === false) parts.push("contabiliza: false");
      return `  ${JSON.stringify(nf)}: { ${parts.join(", ")} }`;
    });

  return `const FATURAMENTO_FINANCEIRO_2026 = Object.freeze({\n${lines.join(",\n")}\n});`;
}

function formatClosedMonths(months) {
  const values = Array.from(months).sort();
  return `const MESES_FINANCEIRO_FECHADOS_2026 = new Set([\n${values.map(month => `  ${JSON.stringify(month)}`).join(",\n")}\n]);`;
}

function buildExpected(locked, closedMonths) {
  return Array.from(closedMonths).sort().reduce((acc, month) => {
    const docs = Object.entries(locked)
      .map(([nf, doc]) => ({ nf: normalizeDigits(nf), doc }))
      .filter(({ doc }) => getMonthKey(doc && doc.data) === month)
      .filter(({ doc }) => !(doc && doc.contabiliza === false))
      .sort((a, b) => a.nf.localeCompare(b.nf, "pt-BR", { numeric: true }));

    acc[month] = {
      total: Math.round(docs.reduce((sum, { doc }) => sum + Number(doc.valor || 0), 0) * 10000) / 10000,
      nfs: docs.map(({ nf }) => nf)
    };
    return acc;
  }, {});
}

function formatExpectedObject(expected) {
  const lines = Object.entries(expected).map(([month, data]) => {
    return `  ${JSON.stringify(month)}: {\n    total: ${formatNumber(data.total)},\n    nfs: ${JSON.stringify(data.nfs)}\n  }`;
  });
  return `const expected = {\n${lines.join(",\n")}\n};`;
}

function replaceBlock(source, pattern, replacement, label) {
  if (!pattern.test(source)) fail(`Nao foi possivel atualizar ${label}.`);
  return source.replace(pattern, replacement);
}

function updateBackend(source, locked, closedMonths) {
  let next = replaceBlock(
    source,
    /const FATURAMENTO_FINANCEIRO_2026 = Object\.freeze\(\{[\s\S]*?\n\}\);/,
    formatFinanceObject(locked),
    "FATURAMENTO_FINANCEIRO_2026"
  );

  next = replaceBlock(
    next,
    /const MESES_FINANCEIRO_FECHADOS_2026 = new Set\(\[[\s\S]*?\n\]\);/,
    formatClosedMonths(closedMonths),
    "MESES_FINANCEIRO_FECHADOS_2026"
  );

  return next;
}

function updateCheckScript(source, expected) {
  return replaceBlock(
    source,
    /const expected = \{[\s\S]*?\n\};/,
    formatExpectedObject(expected),
    "expected do check financeiro"
  );
}

const args = parseArgs(process.argv.slice(2));
const targetMonth = normalizeMonth(args.month || getPreviousMonthKey());
const dryRun = Boolean(args["dry-run"] || args.dryRun);
const force = Boolean(args.force);
const includeList = args.include ? splitList(args.include) : null;
const excludeList = splitList(args.exclude);

const finance = loadFinanceContext();

if (finance.closedMonths.has(targetMonth) && !force) {
  fail(`${targetMonth} ja esta fechado. Use --force apenas se precisar refazer conscientemente o snapshot.`);
}

const response = await fetch(finance.url);
if (!response.ok) fail(`API ERP respondeu ${response.status}.`);

const rows = await response.json();
if (!Array.isArray(rows)) fail("API ERP nao retornou uma lista.");

const groups = new Map();
rows.forEach(row => addRowToGroup(groups, row, targetMonth, includeList, excludeList));

const docs = buildDocsFromGroups(groups);
if (!docs.length) fail(`Nenhuma NF com p_total foi encontrada para ${targetMonth}.`);

const locked = { ...finance.locked };
if (force) {
  Object.entries(locked).forEach(([nf, doc]) => {
    if (getMonthKey(doc && doc.data) === targetMonth) delete locked[nf];
  });
}

docs.forEach(({ nf, doc }) => {
  locked[nf] = doc;
});

const closedMonths = new Set(finance.closedMonths);
closedMonths.add(targetMonth);
const expected = buildExpected(locked, closedMonths);

const totalContabilizado = docs.reduce((sum, { doc }) => {
  return doc.contabiliza === false ? sum : sum + Number(doc.valor || 0);
}, 0);
const nfsContabilizadas = docs.filter(({ doc }) => doc.contabiliza !== false).map(({ nf }) => nf);

console.log(`[finance-lock] Mes: ${targetMonth}`);
console.log(`[finance-lock] NFs contabilizadas: ${nfsContabilizadas.join(", ") || "-"}`);
console.log(`[finance-lock] Total contabilizado: ${formatNumber(totalContabilizado)}`);

if (dryRun) {
  console.log("[finance-lock] Dry-run: nenhum arquivo foi alterado.");
} else {
  writeFileSync(backendPath, updateBackend(finance.source, locked, closedMonths), "utf8");
  writeFileSync(checkPath, updateCheckScript(readFileSync(checkPath, "utf8"), expected), "utf8");

  console.log("[finance-lock] Snapshot financeiro atualizado.");
  console.log("[finance-lock] Rode npm run check para validar a trava.");
}
