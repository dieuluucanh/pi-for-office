/**
 * powerpoint_get_capabilities — report which PowerPoint APIs the current host
 * supports.
 *
 * Read-only: probes `Office.context` requirement sets and never touches the
 * presentation. Host-level, so it is deliberately not wrapped in a
 * `PowerPoint.run` batch.
 */

import {
  POWERPOINT_GET_CAPABILITIES_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  formatCapabilityReport,
  probeCapabilities,
} from "../../office/capabilities.js";
import { t } from "../../language/index.js";

const schema = POWERPOINT_GET_CAPABILITIES_PARAMETERS;
type Params = Static<typeof schema>;

export function createPowerPointGetCapabilitiesTool(): AgentTool<
  typeof schema
> {
  return {
    name: "powerpoint_get_capabilities",
    label: t("tools.powerpointGetCapabilities"),
    description: officeToolDescription("powerpoint.get_capabilities"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      return {
        content: [
          {
            type: "text",
            text: formatCapabilityReport(probeCapabilities("powerpoint")),
          },
        ],
        details: undefined,
      };
    },
  };
}
