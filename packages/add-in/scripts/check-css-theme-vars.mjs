import { promises as fs } from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const THEME_ROOT = path.join(ROOT, "src", "ui", "theme");
const ENTRY_FILE = path.join(ROOT, "src", "ui", "theme.css");

const CSS_VAR_DEF_RE = /(--[A-Za-z0-9_-]+)\s*:/g;
const CSS_VAR_USE_RE = /var\(\s*(--[A-Za-z0-9_-]+)\s*(,[^)]+)?\)/g;

function stripBlockComments(css) {
  return css.replace(/\/\*[\s\S]*?\*\//g, (match) => match.replace(/[^\n]/g, " "));
}

async function collectCssFilesRecursively(dirPath) {
  const entries = await fs.readdir(dirPath, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      const nested = await collectCssFilesRecursively(fullPath);
      files.push(...nested);
      continue;
    }

    if (entry.isFile() && entry.name.endsWith(".css")) {
      files.push(fullPath);
    }
  }

  return files;
}

async function listThemeFiles() {
  const files = await collectCssFilesRecursively(THEME_ROOT);
  files.sort();
  return [ENTRY_FILE, ...files];
}

function toRelative(filePath) {
  return path.relative(ROOT, filePath);
}

function isAllowedExternalVariable(varName) {
  return varName.startsWith("--tw-");
}

async function main() {
  const files = await listThemeFiles();

  const definitions = new Set();
  const usages = [];

  for (const filePath of files) {
    const raw = await fs.readFile(filePath, "utf8");
    const css = stripBlockComments(raw);
    const lines = css.split("\n");

    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index];
      const lineNo = index + 1;

      for (const def of line.matchAll(CSS_VAR_DEF_RE)) {
        definitions.add(def[1]);
      }

      for (const usage of line.matchAll(CSS_VAR_USE_RE)) {
        usages.push({
          filePath,
          lineNo,
          varName: usage[1],
          hasFallback: usage[2] !== undefined,
        });
      }
    }
  }

  const missingHard = usages.filter((usage) => {
    if (definitions.has(usage.varName)) return false;
    if (usage.hasFallback) return false;
    if (isAllowedExternalVariable(usage.varName)) return false;
    return true;
  });

  if (missingHard.length > 0) {
    console.error("\n✗ Undefined theme CSS variables (without fallback):\n");
    for (const miss of missingHard) {
      console.error(`- ${toRelative(miss.filePath)}:${miss.lineNo} → ${miss.varName}`);
    }
    console.error("\nAdd missing tokens to src/ui/theme/tokens.css or provide a fallback in var(...).\n");
    process.exitCode = 1;
    return;
  }

  const uniqueDefs = definitions.size;
  const uniqueUsed = new Set(usages.map((usage) => usage.varName)).size;
  console.log(`✓ Theme CSS variable check passed (${uniqueDefs} defined, ${uniqueUsed} referenced).`);
}

void main();
