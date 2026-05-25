import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");
const distDir = path.join(projectRoot, "dist");

const requiredPaths = [
  "index.html",
  "assents/css/style.css",
  "assents/js/script.js",
  "assents/motor/motorbackend.js",
  "assents/motor/motorcompras.js"
];

function fail(message) {
  console.error(`[build] ${message}`);
  process.exit(1);
}

function assertInsideProject(targetPath) {
  const relative = path.relative(projectRoot, targetPath);
  if (relative.startsWith("..") || path.isAbsolute(relative)) {
    fail(`Caminho fora do projeto bloqueado: ${targetPath}`);
  }
}

function copyRecursive(sourcePath, targetPath) {
  const sourceStats = statSync(sourcePath);

  if (sourceStats.isDirectory()) {
    mkdirSync(targetPath, { recursive: true });
    for (const entry of readdirSync(sourcePath, { withFileTypes: true })) {
      copyRecursive(path.join(sourcePath, entry.name), path.join(targetPath, entry.name));
    }
    return;
  }

  if (sourceStats.isFile()) {
    mkdirSync(path.dirname(targetPath), { recursive: true });
    copyFileSync(sourcePath, targetPath);
  }
}

for (const relativePath of requiredPaths) {
  const absolutePath = path.join(projectRoot, relativePath);
  if (!existsSync(absolutePath)) fail(`Arquivo obrigatorio ausente: ${relativePath}`);
  if (!statSync(absolutePath).isFile()) fail(`Caminho obrigatorio nao e arquivo: ${relativePath}`);
}

const syntax = spawnSync(process.execPath, ["tools/check-syntax.mjs"], {
  cwd: projectRoot,
  stdio: "inherit"
});

if (syntax.status !== 0) {
  fail("Build interrompido por erro de sintaxe.");
}

const financeLock = spawnSync(process.execPath, ["tools/check-finance-lock.mjs"], {
  cwd: projectRoot,
  stdio: "inherit"
});

if (financeLock.status !== 0) {
  fail("Build interrompido por divergencia no snapshot financeiro fechado.");
}

const domainRules = spawnSync(process.execPath, ["tools/check-domain-rules.mjs"], {
  cwd: projectRoot,
  stdio: "inherit"
});

if (domainRules.status !== 0) {
  fail("Build interrompido por divergencia nas regras de obra/pedido.");
}

const indexHtml = readFileSync(path.join(projectRoot, "index.html"), "utf8");
const localAssets = [
  "assents/css/style.css",
  "assents/motor/motorbackend.js",
  "assents/motor/motorcompras.js",
  "assents/js/script.js"
];

for (const asset of localAssets) {
  const assetWithoutQuery = asset.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`${assetWithoutQuery}(?:\\?v=\\d+)?`);
  if (!pattern.test(indexHtml)) {
    fail(`Referencia local esperada nao encontrada no index.html: ${asset}`);
  }
}

assertInsideProject(distDir);
rmSync(distDir, { recursive: true, force: true });
mkdirSync(distDir, { recursive: true });

copyRecursive(path.join(projectRoot, "index.html"), path.join(distDir, "index.html"));
copyRecursive(path.join(projectRoot, "assents"), path.join(distDir, "assents"));

console.log("[build] Build estatico gerado em dist/.");
console.log("[build] Use npm run preview para servir a versao gerada.");
