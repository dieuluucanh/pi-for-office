/**
 * word_insert_hyperlink — Insert a clickable hyperlink with display text into
 * the live Word document.
 *
 * Two paths, because the native API is desktop-only:
 *  - `range.hyperlinks.add()` when WordApiDesktop 1.3 is available;
 *  - otherwise an OOXML `w:fldSimple` HYPERLINK field inserted with
 *    `insertOoxml` (WordApi 1.1), which works on Word for the web too.
 *
 * The OOXML form carries the URL inside the field instruction, so it needs no
 * relationship part and is self-contained — that is what makes it portable.
 */

import {
  WORD_INSERT_HYPERLINK_PARAMETERS,
  officeToolDescription,
} from "@dieulc/pi-office-protocol/office-catalog";
import type { Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordHyperlinkNativeAvailable, wordRun } from "../../word/helpers.js";
import { getErrorMessage } from "../../utils/errors.js";
import { t } from "../../language/index.js";

const schema = WORD_INSERT_HYPERLINK_PARAMETERS;
type Params = Static<typeof schema>;

/** Escape a string for XML text nodes and attribute values. */
function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Build a self-contained OOXML hyperlink field.
 *
 * `w:fldSimple` is a paragraph-level field, so the fragment is a single `w:p`
 * holding the field and a run with the display text. The URL travels inside the
 * `w:instr` attribute, which is why no relationship part is required.
 */
export function buildHyperlinkOoxml(
  text: string,
  url: string,
  screenTip?: string,
): string {
  // HYPERLINK field syntax; `\o` carries the optional screen tip. The raw
  // instruction is built with real quotes and then XML-escaped as a whole, so
  // quotes inside the URL cannot break out of the attribute.
  let instruction = ` HYPERLINK "${url}"`;
  if (screenTip !== undefined) instruction += ` \\o "${screenTip}"`;
  instruction += " ";
  return (
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:fldSimple w:instr="${escapeXml(instruction)}">` +
    `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>` +
    "</w:fldSimple>" +
    "</w:p>"
  );
}

export function createWordInsertHyperlinkTool(): AgentTool<typeof schema> {
  return {
    name: "word_insert_hyperlink",
    label: t("tools.wordInsertHyperlink"),
    description: officeToolDescription("word.insert_hyperlink"),
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      if (params.text.length === 0) {
        return {
          content: [{ type: "text", text: "Missing display text." }],
          details: undefined,
        };
      }
      if (
        !/^https?:\/\//i.test(params.url) &&
        !params.url.startsWith("mailto:")
      ) {
        return {
          content: [
            {
              type: "text",
              text: "url must start with http://, https://, or mailto:.",
            },
          ],
          details: undefined,
        };
      }

      try {
        const location = params.location === "start" ? "Start" : "End";
        const usedNative = await wordRun(async (context) => {
          const body = context.document.body;
          if (wordHyperlinkNativeAvailable()) {
            const range = body.insertText(params.text, location);
            range.hyperlinks.add(range, {
              address: params.url,
              ...(params.screenTip === undefined
                ? {}
                : { screenTip: params.screenTip }),
            });
            await context.sync();
            return true;
          }
          // No WordApiDesktop 1.3 (e.g. Word for the web): the flat field
          // fragment inserts on WordApi 1.1 and needs no relationships.
          body.insertOoxml(
            buildHyperlinkOoxml(params.text, params.url, params.screenTip),
            location,
          );
          await context.sync();
          return false;
        });

        const how = usedNative
          ? "native hyperlink API"
          : "OOXML fallback — this host lacks WordApiDesktop 1.3";
        return {
          content: [
            {
              type: "text",
              text: `Inserted a hyperlink "${params.text}" → ${params.url} at the ${params.location ?? "end"} of the document (${how}).`,
            },
          ],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [
            {
              type: "text",
              text: `Error inserting hyperlink: ${getErrorMessage(e)}`,
            },
          ],
          details: undefined,
        };
      }
    },
  };
}
