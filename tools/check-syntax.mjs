import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");

const files = [
  "assents/js/script.js",
  "assents/motor/motorbackend.js",
  "assents/motor/motorcompras.js"
];

let hasError = false;

for (const relativeFile of files) {
  const absoluteFile = path.join(projectRoot, relativeFile);

  if (!existsSync(absoluteFile)) {
    console.error(`[check] Arquivo nao encontrado: ${relativeFile}`);
    hasError = true;
    continue;
  }

  const result = spawnSync(process.execPath, ["--check", absoluteFile], {
    cwd: projectRoot,
    encoding: "utf8"
  });

  if (result.status !== 0) {
    hasError = true;
    console.error(`[check] Falha de sintaxe em ${relativeFile}`);
    if (result.stdout) process.stdout.write(result.stdout);
    if (result.stderr) process.stderr.write(result.stderr);
  } else {
    console.log(`[check] OK: ${relativeFile}`);
  }
}

if (hasError) {
  process.exitCode = 1;
} else {
  console.log("[check] Validacao de sintaxe concluida.");
}
