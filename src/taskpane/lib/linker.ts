/**
 * linker.ts
 *
 * Run splitting and binding creation for both whole-shape and inline linking.
 */

import type { Binding, Variable } from "../../types";

// ─────────────────────────────────────────────────────────────────────────────
// UUID helper (crypto.randomUUID when available, fallback otherwise)
// ─────────────────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────────────────
// Types used within linker
// ─────────────────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────────────────
// Whole-shape linking
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Create a binding for a whole-shape link.
 * The shape's entire text content is replaced by the variable value.
 */
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

        // Capture original font color before writing
        const textRange = shape.textFrame.textRange;
        textRange.load("text");
        await context.sync();

        // Write the variable value
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

// ─────────────────────────────────────────────────────────────────────────────
// Inline / substring linking
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Split a paragraph's runs at the boundaries of a text selection so the selected
 * substring becomes its own discrete run, then create a binding for it.
 *
 * `selectionStart` and `selectionEnd` are character offsets within the paragraph.
 *
 * NOTE: This function uses the PowerPoint Office JS API to manipulate runs.
 * Validate behavior with Script Lab before Phase 3 deployment (see NOTES.md).
 */
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

        // Build a flat character map: for each char position, which run contains it
        let charPos = 0;
        interface RunSpan {
          runIdx: number;
          start: number;
          end: number;
          text: string;
        }
        const spans: RunSpan[] = [];
        for (let i = 0; i < runs.items.length; i++) {
          const run = runs.items[i];
          run.textRange.load("text");
          await context.sync();
          const text = run.textRange.text;
          spans.push({ runIdx: i, start: charPos, end: charPos + text.length, text });
          charPos += text.length;
        }

        // Identify which runs overlap with [selectionStart, selectionEnd)
        const overlapping = spans.filter(
          (s) => s.end > selectionStart && s.start < selectionEnd
        );
        if (overlapping.length === 0) throw new Error("Selection does not overlap any run");

        // Collect all runs, rebuild text with the selection isolated into a single run.
        // Strategy: delete existing runs and re-insert with the selected span isolated.
        // PowerPoint does not provide a direct "split run" API, so we:
        //   1. Collect the full paragraph text and font properties of each run
        //   2. Rewrite the paragraph with three segments: before, selected, after
        //   3. Apply original formatting to each segment

        const paraTextRange = paragraph.textRange;
        paraTextRange.load("text");
        await context.sync();
        const fullText = paraTextRange.text;

        const before = fullText.slice(0, selectionStart);
        const selected = fullText.slice(selectionStart, selectionEnd);
        const after = fullText.slice(selectionEnd);

        if (!selected) throw new Error("Empty selection");

        // Clear paragraph text and write three runs
        paraTextRange.text = before + selected + after;
        await context.sync();

        // The paragraph now has its text reset; runs are recreated by PowerPoint.
        // We track the run that contains `selected` by character offset.
        // Because we wrote before+selected+after as plain text, PowerPoint may
        // merge them into a single run. We use getSubstring to target the range.

        // Use paragraph.textRange.getSubstring to create a targeted range for the
        // selected text and apply value
        const selectedRange = paraTextRange.getSubstring(selectionStart, selectionEnd - selectionStart);
        selectedRange.text = variable.value;
        await context.sync();

        // Capture the run index after the rewrite
        // Since runs may have been re-created, reload
        const freshRuns = paragraph.textRange.runs;
        freshRuns.load("items");
        await context.sync();

        // Find the run whose text equals variable.value at charOffset
        let targetRunIndex = 0;
        let cumChar = 0;
        for (let i = 0; i < freshRuns.items.length; i++) {
          freshRuns.items[i].textRange.load("text");
          await context.sync();
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

// ─────────────────────────────────────────────────────────────────────────────
// Utility — validate that a shape supports text linking
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns false if the shape type should not be linked (charts, images, etc.).
 * Accepts a PowerPoint shape type string as returned by shape.type.
 */
export function isLinkableShapeType(shapeType: string): boolean {
  const unsupported = ["Chart", "Picture", "SmartArt", "Table", "Media", "3DModel"];
  return !unsupported.includes(shapeType);
}
