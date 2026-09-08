import { promises as fs } from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const BUILTINS_ROOT = path.join(ROOT, "src", "commands", "builtins");

const FORBIDDEN_PATTERNS = [
  /\bstyle\.cssText\b/g,
  /\.style\.[A-Za-z_$][A-Za-z0-9_$]*/g,
  /setAttribute\(\s*["']style["']/g,
];

function toRelative(filePath) {
  return path.relative(ROOT, filePath);
}

async function collectTsFilesRecursively(dirPath) {
  const entries = await fs.readdir(dirPath, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      const nested = await collectTsFilesRecursively(fullPath);
      files.push(...nested);
      continue;
    }

    if (entry.isFile() && entry.name.endsWith(".ts")) {
      files.push(fullPath);
    }
  }

  return files;
}

async function main() {
  const files = await collectTsFilesRecursively(BUILTINS_ROOT);
  files.sort();

  const violations = [];

  for (const filePath of files) {
    const source = await fs.readFile(filePath, "utf8");
    const lines = source.split("\n");

    for (let lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
      const line = lines[lineIndex];

      for (const pattern of FORBIDDEN_PATTERNS) {
        pattern.lastIndex = 0;
        const match = pattern.exec(line);
        if (!match) continue;

        violations.push({
          filePath,
          lineNo: lineIndex + 1,
          snippet: line.trim(),
        });
      }
    }
  }

  if (violations.length > 0) {
    console.error("\n✗ Inline style usage is forbidden in src/commands/builtins/**\n");
    for (const violation of violations) {
      console.error(`- ${toRelative(violation.filePath)}:${violation.lineNo}`);
      console.error(`  ${violation.snippet}`);
    }
    console.error("\nMove styles to src/ui/theme/overlays/*.css and use classes instead.\n");
    process.exitCode = 1;
    return;
  }

  console.log(`✓ Builtins inline-style check passed (${files.length} files).`);
}

void main();
