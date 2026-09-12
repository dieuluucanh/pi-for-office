/**
 * word_insert_image — Insert an inline image from a base64 string (or data
 * URL) into the live Word document. Supports optional width/height (points)
 * and alignment.
 */

import { WORD_INSERT_IMAGE_PARAMETERS } from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const MAX_DECODED_BYTES = 5 * 1024 * 1024; // ~5 MB

const schema = WORD_INSERT_IMAGE_PARAMETERS;
type Params = Static<typeof schema>;

/** Strip a data:...;base64, prefix and validate the base64 payload. */
function normalizeBase64(raw: string): { base64: string; error?: string } {
  const withoutPrefix = raw.replace(/^data:[^;]*;base64,/u, "");
  if (withoutPrefix.length === 0) {
    return { base64: "", error: "Image payload is empty." };
  }
  const pattern = /^[A-Za-z0-9+/=\s]+$/u;
  if (!pattern.test(withoutPrefix)) {
    return {
      base64: "",
      error: "Image payload is not valid base64 (expected [A-Za-z0-9+/=]).",
    };
  }
  const compact = withoutPrefix.replace(/\s+/gu, "");
  const padding = compact.endsWith("==") ? 2 : compact.endsWith("=") ? 1 : 0;
  const decodedBytes = Math.floor((compact.length * 3) / 4) - padding;
  if (decodedBytes > MAX_DECODED_BYTES) {
    return {
      base64: "",
      error: `Image too large (~${Math.round(decodedBytes / 1024 / 1024)} MB); max ${MAX_DECODED_BYTES / 1024 / 1024} MB.`,
    };
  }
  return { base64: compact };
}

export function createWordInsertImageTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_image",
    label: t("tools.wordInsertImage"),
    description:
      "Insert an inline image into the live Word document from a base64 string (or a data URL). " +
      "Max decoded size ~5 MB. To embed a local file, encode it first: on macOS/Linux run `base64 -w0 <file>` " +
      "in bash and pass the output; on Windows run `certutil -encode <file> tmp.b64` and read the file. " +
      "Optionally set width/height in points and paragraph alignment.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      const normalized = normalizeBase64(params.base64);
      if (normalized.error !== undefined) {
        return {
          content: [{ type: "text", text: `Error: ${normalized.error}` }],
          details: undefined,
        };
      }

      try {
        const location = params.location === "start" ? "Start" : "End";
        await wordRun(async (context) => {
          const picture = context.document.body.insertInlinePictureFromBase64(
            normalized.base64,
            location,
          );
          if (params.width !== undefined) picture.width = params.width;
          if (params.height !== undefined) picture.height = params.height;
          await context.sync();
        });

        const applied: string[] = [];
        if (params.width !== undefined) applied.push(`width=${params.width}`);
        if (params.height !== undefined)
          applied.push(`height=${params.height}`);
        return {
          content: [
            {
              type: "text",
              text:
                `Inserted an image at the ${params.location ?? "end"} of the document` +
                (applied.length > 0 ? ` (${applied.join(", ")})` : "") +
                ".",
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting image: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
