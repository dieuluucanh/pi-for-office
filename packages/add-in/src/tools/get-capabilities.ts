/**
 * excel_get_capabilities — report which Excel APIs the current host supports.
 *
 * Read-only: probes `Office.context` requirement sets and never touches the
 * workbook. Like the Word equivalent this is host-level, so it is deliberately
 * not wrapped in `Excel.run` — it still answers when no workbook is attached.
 */

import {
  EXCEL_GET_CAPABILITIES_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  formatCapabilityReport,
  probeCapabilities,
} from "../office/capabilities.js";
import { t } from "../language/index.js";

const schema = EXCEL_GET_CAPABILITIES_PARAMETERS;
type Params = Static<typeof schema>;

export function createExcelGetCapabilitiesTool(): AgentTool<typeof schema> {
  return {
    name: "excel_get_capabilities",
    label: t("tools.excelGetCapabilities"),
    description: officeToolDescription("excel.get_capabilities"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      return {
        content: [
          {
            type: "text",
            text: formatCapabilityReport(probeCapabilities("excel")),
          },
        ],
        details: undefined,
      };
    },
  };
}
