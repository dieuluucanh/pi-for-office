/** SettingsStore-backed session/workbook association adapter shared by hosts. */

import type { SpreadsheetHostSessionStorage } from "./types.js";
import {
  clearSessionWorkbookPending,
  forgetSessionWorkbook,
  getLatestSessionForWorkbook,
  getSessionWorkbookId,
  linkSessionToWorkbook,
  markSessionWorkbookPending,
  migrateAliasesToCanonical,
  partitionSessionIdsByWorkbook,
  readSessionWorkbookLinks,
  relinkSessionToWorkbook,
  setLatestSessionForWorkbook,
} from "../workbook/session-association.js";

export const settingsBackedSessionStorage: SpreadsheetHostSessionStorage = {
  getSessionWorkbookId,
  linkSessionToWorkbook,
  relinkSessionToWorkbook,
  forgetSessionWorkbook,
  setLatestSessionForWorkbook,
  getLatestSessionForWorkbook,
  readSessionWorkbookLinks,
  migrateAliasesToCanonical,
  markSessionWorkbookPending,
  clearSessionWorkbookPending,
  partitionSessionIdsByWorkbook,
};
