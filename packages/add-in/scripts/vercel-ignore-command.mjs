import { pathToFileURL } from "node:url";

const BUILD_EXIT_CODE = 1;
const SKIP_EXIT_CODE = 0;

function normalizeEnvValue(value) {
  return typeof value === "string" ? value.trim() : "";
}

export function resolveVercelIgnoreCommandExitCode(environment) {
  const commitRef = normalizeEnvValue(environment.VERCEL_GIT_COMMIT_REF);
  const pullRequestId = normalizeEnvValue(environment.VERCEL_GIT_PULL_REQUEST_ID);

  if (commitRef.length === 0 || commitRef === "main" || pullRequestId.length > 0) {
    return BUILD_EXIT_CODE;
  }

  return SKIP_EXIT_CODE;
}

function isCurrentModuleEntrypoint() {
  const entrypoint = process.argv[1];
  if (typeof entrypoint !== "string") {
    return false;
  }

  return pathToFileURL(entrypoint).href === import.meta.url;
}

if (isCurrentModuleEntrypoint()) {
  const exitCode = resolveVercelIgnoreCommandExitCode(process.env);

  if (exitCode === SKIP_EXIT_CODE) {
    const commitRef = normalizeEnvValue(process.env.VERCEL_GIT_COMMIT_REF);
    console.log(`Skipping non-PR branch deploy for ${commitRef}`);
  }

  process.exit(exitCode);
}
