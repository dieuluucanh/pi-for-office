/**
 * word_get_capabilities — report which Word APIs the current host supports.
 *
 * Read-only: probes `Office.context` requirement sets and never touches the
 * document. This is what lets the agent answer "can this host do X?" from the
 * real host instead of guessing — the failure that produced a false "Word has
 * no text highlight" answer.
 *
 * Deliberately not wrapped in `Word.run`: capability probing is host-level, so
 * it still answers when no document batch is available.
 */

import {
  WORD_GET_CAPABILITIES_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  formatCapabilityReport,
  probeWordCapabilities,
} from "../../office/capabilities.js";
import { t } from "../../language/index.js";

const schema = WORD_GET_CAPABILITIES_PARAMETERS;
type Params = Static<typeof schema>;

export function createWordGetCapabilitiesTool(): AgentTool<typeof schema> {
  return {
    name: "word_get_capabilities",
    label: t("tools.wordGetCapabilities"),
    description: officeToolDescription("word.get_capabilities"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      return {
        content: [
          {
            type: "text",
            text: formatCapabilityReport(probeWordCapabilities()),
          },
        ],
        details: undefined,
      };
    },
  };
}
