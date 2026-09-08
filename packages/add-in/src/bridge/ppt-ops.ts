/**
 * PowerPoint bridge executors — first PowerPoint.js operations in pi-for-office.
 *
 * Runs inside the PowerPoint task pane via the `PowerPoint.run` batch model.
 * Note: slide size / precise geometry depend on newer API requirement sets;
 * ops degrade with clear messages when a shape API is unsupported.
 */

import { guardExecutor, type OfficeOpExecutor } from "./ops.js";

/** Read the text of every text-bearing shape on a slide. */
async function readSlideTexts(slide: PowerPoint.Slide): Promise<
  Array<{ name: string; type: string; text: string }>
> {
  // Phase 1: shape proxies + which ones carry text.
  const shapes = slide.shapes;
  shapes.load("items/type/name");
  // textFrame presence must be probed separately to avoid touching text on
  // non-text shapes.
  shapes.load("items/textFrame/hasText");
  await shapes.context.sync();

  // Phase 2: load text for text-bearing shapes.
  const texts = shapes.items.map((shape) => {
    if (shape.textFrame && shape.textFrame.hasText) {
      const range = shape.textFrame.textRange;
      range.load("text");
      return { shape, range };
    }
    return null;
  });
  await shapes.context.sync();

  const out: Array<{ name: string; type: string; text: string }> = [];
  for (const entry of texts) {
    if (entry === null) continue;
    const text = entry.range.text ?? "";
    if (text.trim().length === 0) continue;
    out.push({ name: entry.shape.name, type: entry.shape.type, text });
  }
  return out;
}

export const POWERPOINT_OPS: ReadonlyMap<string, OfficeOpExecutor> = new Map<
  string,
  OfficeOpExecutor
>([
  [
    "powerpoint.get_overview",
    guardExecutor(async () => {
      const result = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const perSlide = slides.items.map((slide) => {
          slide.load("shapes");
          return slide;
        });
        await context.sync();

        const counts = perSlide.map((slide) => ({
          id: slide.id,
          shapeCount: slide.shapes.items.length,
        }));
        return counts;
      });

      const lines: string[] = [`**Presentation overview**`];
      lines.push(`- Slides: ${result.length}`);
      for (let i = 0; i < result.length; i += 1) {
        const slide = result[i];
        if (slide === undefined) continue;
        lines.push(`  - Slide ${i + 1}: ${slide.shapeCount} shapes`);
      }
      lines.push("Use powerpoint.read_slide to read a slide's text; powerpoint.add_slide / add_text_box to edit.");

      return { text: lines.join("\n"), details: { slideCount: result.length, slides: result } };
    }),
  ],
  [
    "powerpoint.read_slide",
    guardExecutor(async (args) => {
      const rawIndex = args.slideIndex;
      const slideIndex = typeof rawIndex === "number" ? Math.floor(rawIndex) : NaN;
      if (!Number.isFinite(slideIndex) || slideIndex < 1) {
        return { text: "slideIndex must be a 1-based slide number.", isError: true };
      }

      const out = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (slides.items.length < slideIndex) {
          return { error: `Presentation has only ${slides.items.length} slide(s).` };
        }
        const slide = slides.items[slideIndex - 1];
        if (slide === undefined) {
          return { error: "Slide not found." };
        }
        const slideTexts = await readSlideTexts(slide);
        return { slideTexts, id: slide.id };
      });

      if ("error" in out) {
        return { text: out.error, isError: true };
      }

      const lines: string[] = [`**Slide ${slideIndex}** (id: ${out.id})`];
      if (out.slideTexts.length === 0) {
        lines.push("_No text content on this slide._");
      } else {
        for (const item of out.slideTexts) {
          lines.push("");
          lines.push(`**${item.name || item.type}:**`);
          lines.push(item.text);
        }
      }
      return { text: lines.join("\n"), details: { slideIndex, texts: out.slideTexts } };
    }),
  ],
  [
    "powerpoint.add_slide",
    guardExecutor(async () => {
      const newId = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        const countResult = slides.getCount();
        await context.sync();
        slides.add();
        await context.sync();
        // The appended slide sits at index countResult.value.
        const added = slides.getItemAt(countResult.value);
        added.load("id");
        await context.sync();
        context.presentation.setSelectedSlides([added.id]);
        await context.sync();
        return added.id;
      });
      return { text: `Added a new slide (id: ${newId}) and navigated to it.`, details: { id: newId } };
    }),
  ],
  [
    "powerpoint.add_text_box",
    guardExecutor(async (args) => {
      const slideIndex = typeof args.slideIndex === "number" ? Math.floor(args.slideIndex) : NaN;
      const text = typeof args.text === "string" ? args.text : "";
      if (!Number.isFinite(slideIndex) || slideIndex < 1) {
        return { text: "slideIndex must be a 1-based slide number.", isError: true };
      }
      if (text.length === 0) {
        return { text: "Nothing to add (empty text).", isError: true };
      }

      const out = await PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        if (slides.items.length < slideIndex) {
          return { error: `Presentation has only ${slides.items.length} slide(s).` };
        }
        const slide = slides.items[slideIndex - 1];
        if (slide === undefined) return { error: "Slide not found." };

        const options: PowerPoint.ShapeAddOptions = {
          left: typeof args.x === "number" ? args.x : 72,
          top: typeof args.y === "number" ? args.y : 72,
          width: typeof args.width === "number" ? args.width : 400,
          height: typeof args.height === "number" ? args.height : 60,
        };
        const shape = slide.shapes.addTextBox(text, options);
        await context.sync();
        shape.load("id");
        await context.sync();
        return { id: shape.id };
      });

      if ("error" in out) {
        return { text: out.error, isError: true };
      }
      return {
        text: `Added a text box (${text.length} chars) to slide ${slideIndex}.`,
        details: { slideIndex, id: out.id },
      };
    }),
  ],
]);
