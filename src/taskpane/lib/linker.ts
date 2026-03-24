import type { Binding, Variable } from "../../types";

export function generateId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  // Fallback for environments without randomUUID
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

export interface RunInfo {
  text: string;
  index: number;
}

export interface SplitResult {
  /** Index of the newly isolated run that contains the linked text */
  runIndex: number;
  /** Character offset of the split run within its paragraph */
  charOffset: number;
}

export async function linkWholeShape(
  variable: Variable,
  shapeId: string,
  shapeName: string,
  slideIndex: number
): Promise<Binding> {
  return new Promise<Binding>((resolve, reject) => {
    PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[slideIndex];
        if (!slide) throw new Error(`Slide ${slideIndex} not found`);

        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        const shape = shapes.items.find((s) => s.id === shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on slide ${slideIndex}`);

        const textRange = shape.textFrame.textRange;
        textRange.load("text");
        await context.sync();

        textRange.text = variable.value;
        await context.sync();

        const binding: Binding = {
          id: generateId(),
          variableName: variable.name,
          shapeId,
          slideIndex,
          paragraphIndex: 0,
          runIndex: 0,
          lastKnownValue: variable.value,
          charOffset: 0,
          originalFontColor: null,
          originalHighlightColor: null,
        };

        resolve(binding);
      } catch (err) {
        reject(err);
      }
    }).catch(reject);
  });
}

// NOTE: Validate run-level API behavior with Script Lab before Phase 3 deployment (see NOTES.md).
export async function linkInlineSelection(
  variable: Variable,
  shapeId: string,
  slideIndex: number,
  paragraphIndex: number,
  selectionStart: number,
  selectionEnd: number
): Promise<Binding> {
  return new Promise<Binding>((resolve, reject) => {
    PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[slideIndex];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        const shape = shapes.items.find((s) => s.id === shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        const paragraphs = shape.textFrame.textRange.paragraphs;
        paragraphs.load("items");
        await context.sync();

        const paragraph = paragraphs.items[paragraphIndex];
        if (!paragraph) throw new Error(`Paragraph ${paragraphIndex} not found`);

        const runs = paragraph.textRange.runs;
        runs.load("items");
        await context.sync();

        // Build a flat character map. Batch-load all run texts in a single sync.
        for (const run of runs.items) {
          run.textRange.load("text");
        }
        await context.sync();

        interface RunSpan {
          runIdx: number;
          start: number;
          end: number;
          text: string;
        }
        let charPos = 0;
        const spans: RunSpan[] = [];
        for (let i = 0; i < runs.items.length; i++) {
          const text = runs.items[i].textRange.text;
          spans.push({ runIdx: i, start: charPos, end: charPos + text.length, text });
          charPos += text.length;
        }

        const overlapping = spans.filter(
          (s) => s.end > selectionStart && s.start < selectionEnd
        );
        if (overlapping.length === 0) throw new Error("Selection does not overlap any run");

        // PowerPoint has no direct "split run" API. Strategy:
        //   1. Rewrite paragraph text (before + selected + after) — preserves content
        //   2. Use getSubstring to target the selected range and write the variable value
        const paraTextRange = paragraph.textRange;
        paraTextRange.load("text");
        await context.sync();
        const fullText = paraTextRange.text;

        const before = fullText.slice(0, selectionStart);
        const selected = fullText.slice(selectionStart, selectionEnd);
        const after = fullText.slice(selectionEnd);

        if (!selected) throw new Error("Empty selection");

        paraTextRange.text = before + selected + after;
        await context.sync();

        const selectedRange = paraTextRange.getSubstring(selectionStart, selectionEnd - selectionStart);
        selectedRange.text = variable.value;
        await context.sync();

        // Reload runs after rewrite and batch-load their texts in one sync.
        const freshRuns = paragraph.textRange.runs;
        freshRuns.load("items");
        await context.sync();

        for (const run of freshRuns.items) {
          run.textRange.load("text");
        }
        await context.sync();

        let targetRunIndex = 0;
        let cumChar = 0;
        for (let i = 0; i < freshRuns.items.length; i++) {
          const t = freshRuns.items[i].textRange.text;
          if (cumChar === selectionStart && t === variable.value) {
            targetRunIndex = i;
            break;
          }
          cumChar += t.length;
          targetRunIndex = i;
        }

        const binding: Binding = {
          id: generateId(),
          variableName: variable.name,
          shapeId,
          slideIndex,
          paragraphIndex,
          runIndex: targetRunIndex,
          lastKnownValue: variable.value,
          charOffset: selectionStart,
          originalFontColor: null,
          originalHighlightColor: null,
        };

        resolve(binding);
      } catch (err) {
        reject(err);
      }
    }).catch(reject);
  });
}

export function isLinkableShapeType(shapeType: string): boolean {
  const unsupported = ["Chart", "Picture", "SmartArt", "Table", "Media", "3DModel"];
  return !unsupported.includes(shapeType);
}
