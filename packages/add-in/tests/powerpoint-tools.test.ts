/**
 * PowerPoint tool tests — shape fill (step 4).
 *
 * Exercised against the fake `PowerPoint` global. Verifies each property
 * reaches the right Office.js surface and that the requirement gates fail
 * loudly on an older host while still working on a capable one.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

import { POWERPOINT_UNDERLINE_STYLES } from "@dieulc/pi-office-protocol/office-catalog";
import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createPowerPointAddTextBoxTool } from "../src/tools/powerpoint/add-text-box.ts";
import { createPowerPointAddSlideTool } from "../src/tools/powerpoint/add-slide.ts";
import { createPowerPointFormatSlideTool } from "../src/tools/powerpoint/format-slide.ts";
import { createPowerPointReadSlideTool } from "../src/tools/powerpoint/read-slide.ts";
import { createPowerPointGetOverviewTool } from "../src/tools/powerpoint/get-overview.ts";
import {
  installFakeAppStorage,
  makeShape,
  withFakePowerPoint,
} from "./helpers/fake-powerpoint.ts";

installFakeAppStorage();

function toolText(result: AgentToolResult<undefined>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text.");
  }
  return first.text;
}

void test("add_text_box applies a solid fill to the created shape", async () => {
  await withFakePowerPoint(async (context) => {
    const tool = createPowerPointAddTextBoxTool();
    const result = await tool.execute("tc-ppt-fill", {
      slideIndex: 1,
      text: "hello",
      fillColor: "#FFFF00",
      fillTransparency: 0.25,
    });
    assert.match(toolText(result), /\(formatted\)/);

    const shape = context.presentation.slides.items[0]?.shapes.added[0];
    assert.deepEqual(shape?.fill.solidColors, ["#FFFF00"]);
    assert.equal(shape?.fill.transparency, 0.25);
  }, {});
});

void test("add_text_box fillClear clears the fill and a color still wins", async () => {
  await withFakePowerPoint(async (context) => {
    const tool = createPowerPointAddTextBoxTool();
    await tool.execute("tc-ppt-fill-clear", {
      slideIndex: 1,
      text: "hello",
      fillClear: true,
      fillColor: "#00FF00",
    });

    const shape = context.presentation.slides.items[0]?.shapes.added[0];
    assert.equal(shape?.fill.cleared, 1);
    assert.deepEqual(shape?.fill.solidColors, ["#00FF00"]);
  }, {});
});

void test("add_text_box fill fails loudly below PowerPointApi 1.4", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-fill-gate", {
        slideIndex: 1,
        text: "hello",
        fillColor: "#FFFF00",
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.3" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.4/);
  assert.match(text, /platform: /);
});

void test("add_text_box fill works when the host has PowerPointApi 1.4", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddTextBoxTool();
      const result = await tool.execute("tc-ppt-fill-ok", {
        slideIndex: 1,
        text: "hello",
        fillColor: "#123456",
      });
      assert.match(toolText(result), /Added a text box/);
      assert.deepEqual(
        context.presentation.slides.items[0]?.shapes.added[0]?.fill.solidColors,
        ["#123456"],
      );
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );
});

void test("format_slide applies fill to text-bearing shapes only", async () => {
  await withFakePowerPoint(async (context) => {
    const slide = context.presentation.slides.items[0];
    if (!slide) throw new Error("No slide in fake context.");
    const textShape = makeShape({ hasText: true });
    const emptyShape = makeShape({ hasText: false });
    slide.shapes.items.push(textShape, emptyShape);

    const tool = createPowerPointFormatSlideTool();
    const result = await tool.execute("tc-ppt-slide-fill", {
      slideIndex: 1,
      fillColor: "#ABCDEF",
    });

    assert.match(toolText(result), /Formatted 1 text box/);
    assert.deepEqual(textShape.fill.solidColors, ["#ABCDEF"]);
    assert.deepEqual(emptyShape.fill.solidColors, []);
  }, {});
});

void test("format_slide applies a slide background fill (PowerPointApi 1.10)", async () => {
  await withFakePowerPoint(
    async (context) => {
      const slide = context.presentation.slides.items[0];
      if (!slide) throw new Error("No slide in fake context.");
      slide.shapes.items.push(makeShape({ hasText: true }));

      const tool = createPowerPointFormatSlideTool();
      const result = await tool.execute("tc-ppt-bg", {
        slideIndex: 1,
        slideBackgroundColor: "#FF00FF",
        slideBackgroundTransparency: 0.5,
      });

      assert.match(toolText(result), /background=#FF00FF/);
      assert.deepEqual(slide.background.fill.solid.options, [
        { color: "#FF00FF", transparency: 0.5 },
      ]);
    },
    { capabilities: { supported: { PowerPointApi: "1.10" } } },
  );
});

void test("format_slide background fails loudly below PowerPointApi 1.10", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointFormatSlideTool();
      return tool.execute("tc-ppt-bg-gate", {
        slideIndex: 1,
        slideBackgroundColor: "#FF00FF",
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.10/);
  assert.match(text, /slide background fill/);
});

void test("format_slide background clear resets the background", async () => {
  await withFakePowerPoint(
    async (context) => {
      const slide = context.presentation.slides.items[0];
      if (!slide) throw new Error("No slide in fake context.");
      slide.shapes.items.push(makeShape({ hasText: true }));

      const tool = createPowerPointFormatSlideTool();
      await tool.execute("tc-ppt-bg-clear", {
        slideIndex: 1,
        slideBackgroundClear: true,
      });

      assert.equal(slide.background.resetCalls, 1);
    },
    { capabilities: { supported: { PowerPointApi: "1.10" } } },
  );
});

void test("format_slide reports when nothing was requested", async () => {
  const result = await withFakePowerPoint(async () => {
    const tool = createPowerPointFormatSlideTool();
    return tool.execute("tc-ppt-none", { slideIndex: 1 });
  }, {});

  assert.match(toolText(result), /No formatting options provided/);
});

void test("add_text_box applies line color, weight, style and dash style", async () => {
  await withFakePowerPoint(async (context) => {
    const tool = createPowerPointAddTextBoxTool();
    const result = await tool.execute("tc-ppt-line", {
      slideIndex: 1,
      text: "hello",
      lineColor: "#00FF00",
      lineWeight: 2.5,
      lineStyle: "ThickThin",
      lineDashStyle: "LongDashDot",
      lineTransparency: 0.1,
    });
    assert.match(toolText(result), /\(formatted\)/);

    const shape = context.presentation.slides.items[0]?.shapes.added[0];
    assert.equal(shape?.lineFormat.color, "#00FF00");
    assert.equal(shape?.lineFormat.weight, 2.5);
    assert.equal(shape?.lineFormat.style, "ThickThin");
    assert.equal(shape?.lineFormat.dashStyle, "LongDashDot");
    assert.equal(shape?.lineFormat.transparency, 0.1);
  }, {});
});

void test("lineVisible is applied after the other line properties", async () => {
  await withFakePowerPoint(async (context) => {
    const tool = createPowerPointAddTextBoxTool();
    await tool.execute("tc-ppt-line-hidden", {
      slideIndex: 1,
      text: "hello",
      lineColor: "#FF0000",
      lineVisible: false,
    });

    const lineFormat =
      context.presentation.slides.items[0]?.shapes.added[0]?.lineFormat;
    assert.equal(lineFormat?.color, "#FF0000");
    assert.equal(lineFormat?.visible, false);
  }, {});
});

void test("line format fails loudly below PowerPointApi 1.4", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-line-gate", {
        slideIndex: 1,
        text: "hello",
        lineColor: "#000000",
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.3" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.4/);
  assert.match(text, /shape line/);
});

void test("format_slide applies line formatting to text-bearing shapes", async () => {
  await withFakePowerPoint(async (context) => {
    const slide = context.presentation.slides.items[0];
    if (!slide) throw new Error("No slide in fake context.");
    const shape = makeShape({ hasText: true });
    slide.shapes.items.push(shape);

    const tool = createPowerPointFormatSlideTool();
    const result = await tool.execute("tc-ppt-slide-line", {
      slideIndex: 1,
      lineDashStyle: "SquareDot",
    });

    assert.match(toolText(result), /lineDashStyle=SquareDot/);
    assert.equal(shape.lineFormat.dashStyle, "SquareDot");
  }, {});
});

void test("bulletVisible works on a PowerPointApi 1.4 host", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddTextBoxTool();
      const result = await tool.execute("tc-ppt-bullet-visible", {
        slideIndex: 1,
        text: "hello",
        bulletVisible: false,
      });
      assert.match(toolText(result), /\(formatted\)/);

      const bullet =
        context.presentation.slides.items[0]?.shapes.added[0]?.textFrame
          .textRange.paragraphFormat.bulletFormat;
      assert.equal(bullet?.visible, false);
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );
});

void test("bullet style needs PowerPointApi 1.10 even on a 1.4 host", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-bullet-gate", {
        slideIndex: 1,
        text: "hello",
        bulletStyle: "ArabicNumeralPeriod",
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.10/);
  assert.match(text, /bullet style/);
});

void test("indent level needs PowerPointApi 1.10", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-indent-gate", {
        slideIndex: 1,
        text: "hello",
        indentLevel: 2,
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.10/);
  assert.match(text, /paragraph indent level/);
});

void test("bullet style and indent apply on a PowerPointApi 1.10 host", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddTextBoxTool();
      await tool.execute("tc-ppt-bullet-ok", {
        slideIndex: 1,
        text: "hello",
        bulletStyle: "RomanUppercasePeriod",
        bulletType: "Numbered",
        indentLevel: 2,
      });

      const paragraph =
        context.presentation.slides.items[0]?.shapes.added[0]?.textFrame
          .textRange.paragraphFormat;
      assert.equal(paragraph?.bulletFormat.style, "RomanUppercasePeriod");
      assert.equal(paragraph?.bulletFormat.type, "Numbered");
      assert.equal(paragraph?.indentLevel, 2);
    },
    { capabilities: { supported: { PowerPointApi: "1.10" } } },
  );
});

void test("format_slide applies indent and bullets to text-bearing shapes", async () => {
  await withFakePowerPoint(
    async (context) => {
      const slide = context.presentation.slides.items[0];
      if (!slide) throw new Error("No slide in fake context.");
      const shape = makeShape({ hasText: true });
      slide.shapes.items.push(shape);

      const tool = createPowerPointFormatSlideTool();
      const result = await tool.execute("tc-ppt-slide-bullets", {
        slideIndex: 1,
        bulletVisible: true,
        indentLevel: 1,
      });

      assert.match(toolText(result), /bulletVisible=true/);
      assert.match(toolText(result), /indentLevel=1/);
      assert.equal(
        shape.textFrame.textRange.paragraphFormat.bulletFormat.visible,
        true,
      );
      assert.equal(shape.textFrame.textRange.paragraphFormat.indentLevel, 1);
    },
    { capabilities: { supported: { PowerPointApi: "1.10" } } },
  );
});

void test("underline accepts the legacy boolean alias and a named style", async () => {
  await withFakePowerPoint(async (context) => {
    const tool = createPowerPointAddTextBoxTool();
    await tool.execute("tc-ppt-underline-true", {
      slideIndex: 1,
      text: "a",
      underline: true,
    });
    await tool.execute("tc-ppt-underline-style", {
      slideIndex: 1,
      text: "b",
      underline: "WavyDouble",
    });

    const added = context.presentation.slides.items[0]?.shapes.added ?? [];
    assert.equal(added[0]?.textFrame.textRange.font.underline, "Single");
    assert.equal(added[1]?.textFrame.textRange.font.underline, "WavyDouble");
  }, {});
});

void test("every PowerPoint underline style is reachable from the schema", async () => {
  // The catalog constant is the single source for both the schema union and
  // the runtime resolver.
  assert.equal(POWERPOINT_UNDERLINE_STYLES.length, 17);
  assert.ok(
    (POWERPOINT_UNDERLINE_STYLES as readonly string[]).includes("WavyDouble"),
  );
  assert.ok(
    !(POWERPOINT_UNDERLINE_STYLES as readonly string[]).includes("Hair"),
  );
});

void test("font effects apply on a PowerPointApi 1.8 host", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddTextBoxTool();
      const result = await tool.execute("tc-ppt-effects", {
        slideIndex: 1,
        text: "hello",
        strikethrough: true,
        doubleStrikethrough: false,
        smallCaps: true,
        allCaps: true,
        superscript: true,
      });
      assert.match(toolText(result), /\(formatted\)/);

      const font =
        context.presentation.slides.items[0]?.shapes.added[0]?.textFrame
          .textRange.font;
      assert.equal(font?.strikethrough, true);
      assert.equal(font?.doubleStrikethrough, false);
      assert.equal(font?.smallCaps, true);
      assert.equal(font?.allCaps, true);
      assert.equal(font?.superscript, true);
    },
    { capabilities: { supported: { PowerPointApi: "1.8" } } },
  );
});

void test("font effects fail loudly below PowerPointApi 1.8", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-effects-gate", {
        slideIndex: 1,
        text: "hello",
        smallCaps: true,
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.8/);
  assert.match(text, /small caps/);
});

void test("add_slide navigates when setSelectedSlides is available", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddSlideTool();
      const result = await tool.execute("tc-ppt-slide", {});

      assert.match(toolText(result), /navigated to it/);
      assert.equal(context.presentation.selectionCalls.length, 1);
    },
    { capabilities: { supported: { PowerPointApi: "1.5" } } },
  );
});

void test("add_slide still adds the slide when setSelectedSlides is unsupported", async () => {
  await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddSlideTool();
      const result = await tool.execute("tc-ppt-slide-gate", {});

      assert.match(toolText(result), /navigation skipped/);
      assert.equal(context.presentation.selectionCalls.length, 0);
      assert.equal(context.presentation.slides.items.length, 2);
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );
});

/* ── Shape access (PowerPointApi 1.3) and core text format (1.4) ────────── */

void test("add_text_box fails loudly below PowerPointApi 1.3 (shape access)", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-shapes-gate", { slideIndex: 1, text: "hi" });
    },
    { capabilities: { supported: { PowerPointApi: "1.2" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.3/);
  assert.match(text, /shape access/);
});

void test("format_slide fails loudly below PowerPointApi 1.3 (shape access)", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointFormatSlideTool();
      return tool.execute("tc-ppt-slide-shapes-gate", {
        slideIndex: 1,
        bold: true,
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.2" } } },
  );

  assert.match(toolText(result), /needs PowerPointApi 1\.3/);
});

void test("read_slide fails loudly below PowerPointApi 1.3 and reads on 1.3", async () => {
  const gated = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointReadSlideTool();
      return tool.execute("tc-ppt-read-gate", { slideIndex: 1 });
    },
    { capabilities: { supported: { PowerPointApi: "1.2" } } },
  );
  assert.match(toolText(gated), /needs PowerPointApi 1\.3/);

  const ok = await withFakePowerPoint(
    async (context) => {
      const slide = context.presentation.slides.items[0];
      if (!slide) throw new Error("No slide in fake context.");
      const shape = makeShape({ hasText: true });
      shape.textFrame.textRange.text = "hello world";
      slide.shapes.items.push(shape);

      const tool = createPowerPointReadSlideTool();
      return tool.execute("tc-ppt-read-ok", { slideIndex: 1 });
    },
    { capabilities: { supported: { PowerPointApi: "1.3" } } },
  );
  assert.match(toolText(ok), /hello world/);
});

void test("get_overview fails loudly below PowerPointApi 1.3 and reads on 1.3", async () => {
  const gated = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointGetOverviewTool();
      return tool.execute("tc-ppt-ov-gate", {});
    },
    { capabilities: { supported: { PowerPointApi: "1.2" } } },
  );
  assert.match(toolText(gated), /needs PowerPointApi 1\.3/);

  const ok = await withFakePowerPoint(
    async (context) => {
      const slide = context.presentation.slides.items[0];
      if (!slide) throw new Error("No slide in fake context.");
      slide.shapes.items.push(makeShape({ hasText: true }));

      const tool = createPowerPointGetOverviewTool();
      return tool.execute("tc-ppt-ov-ok", {});
    },
    { capabilities: { supported: { PowerPointApi: "1.3" } } },
  );
  assert.match(toolText(ok), /Slide 1: 1 shapes/);
});

void test("core font formatting fails loudly below PowerPointApi 1.4", async () => {
  const result = await withFakePowerPoint(
    async () => {
      const tool = createPowerPointAddTextBoxTool();
      return tool.execute("tc-ppt-textformat-gate", {
        slideIndex: 1,
        text: "hi",
        bold: true,
      });
    },
    { capabilities: { supported: { PowerPointApi: "1.3" } } },
  );

  const text = toolText(result);
  assert.match(text, /needs PowerPointApi 1\.4/);
  assert.match(text, /text frame, font and paragraph formatting/);
});

void test("core font formatting applies on PowerPointApi 1.4", async () => {
  const bold = await withFakePowerPoint(
    async (context) => {
      const tool = createPowerPointAddTextBoxTool();
      await tool.execute("tc-ppt-textformat-ok", {
        slideIndex: 1,
        text: "hi",
        bold: true,
      });
      return context.presentation.slides.items[0]?.shapes.added[0]?.textFrame
        .textRange.font.bold;
    },
    { capabilities: { supported: { PowerPointApi: "1.4" } } },
  );

  assert.equal(bold, true);
});
