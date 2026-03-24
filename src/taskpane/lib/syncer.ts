/**
 * syncer.ts
 *
 * Validation, classification, and sync of bindings.
 *
 * Binding states:
 *   Clean       — tagged run is intact and contains lastKnownValue
 *   Recoverable — run has been remerged but lastKnownValue found near charOffset
 *   Broken      — lastKnownValue not found in paragraph; requires user action
 */

import type {
  Binding,
  BrokenBinding,
  ClassifiedBinding,
  BindingState,
  VarSyncRegistry,
  SyncSummary,
  Variable,
} from "../../types";

// ─────────────────────────────────────────────────────────────────────────────
// Classification helpers (pure, no Office JS — fully unit-testable)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Determine the state of a single binding given the current paragraph text
 * and an optional list of current run texts.
 *
 * @param binding        The binding to classify
 * @param runTexts       Current text of each run in the paragraph (may be empty if paragraph was rewritten)
 * @param paragraphText  Full text of the paragraph
 */
export function classifyBinding(
  binding: Binding,
  runTexts: string[],
  paragraphText: string
): BindingState {
  const lkv = binding.lastKnownValue;

  // 1. Check if the expected run still contains lastKnownValue
  if (
    runTexts.length > binding.runIndex &&
    runTexts[binding.runIndex] === lkv
  ) {
    return "clean";
  }

  // 2. Check if lastKnownValue appears in the paragraph near charOffset
  if (paragraphText.includes(lkv)) {
    const idx = paragraphText.indexOf(lkv);
    const delta = Math.abs(idx - binding.charOffset);
    // Allow up to half the lkv length as drift tolerance
    if (delta <= Math.max(lkv.length / 2, 5)) {
      return "recoverable";
    }
    // Found elsewhere in paragraph but too far — still recoverable (user probably moved it)
    return "recoverable";
  }

  return "broken";
}

/**
 * Find the character offset of lastKnownValue within paragraphText.
 * Returns -1 if not found.
 */
export function findLastKnownValueOffset(
  lastKnownValue: string,
  paragraphText: string
): number {
  return paragraphText.indexOf(lastKnownValue);
}

// ─────────────────────────────────────────────────────────────────────────────
// Office JS — sync a single variable across all its bindings
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Sync all bindings for the given variable.
 * Returns a SyncSummary describing what happened.
 */
export async function syncVariable(
  variable: Variable,
  bindings: Binding[]
): Promise<{ updatedBindings: Binding[]; summary: SyncSummary }> {
  const variableBindings = bindings.filter((b) => b.variableName === variable.name);

  const classified: ClassifiedBinding[] = [];
  const brokenList: BrokenBinding[] = [];
  const updatedBindings = [...bindings];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const binding of variableBindings) {
      const slide = slides.items[binding.slideIndex];
      if (!slide) {
        brokenList.push({ ...binding, slideNumber: binding.slideIndex + 1, shapeName: binding.shapeId });
        classified.push({ binding, state: "broken" });
        continue;
      }

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      const shape = shapes.items.find((s) => s.id === binding.shapeId);
      if (!shape) {
        brokenList.push({ ...binding, slideNumber: binding.slideIndex + 1, shapeName: binding.shapeId });
        classified.push({ binding, state: "broken" });
        continue;
      }

      shape.textFrame.textRange.load("text");
      shape.load("name");
      await context.sync();

      const shapeName = shape.name;
      const paraRange = shape.textFrame.textRange.paragraphs;
      paraRange.load("items");
      await context.sync();

      const para = paraRange.items[binding.paragraphIndex];
      if (!para) {
        brokenList.push({
          ...binding,
          slideNumber: binding.slideIndex + 1,
          shapeName,
        });
        classified.push({ binding, state: "broken" });
        continue;
      }

      para.textRange.load("text");
      const runs = para.textRange.runs;
      runs.load("items");
      await context.sync();

      const runTexts: string[] = [];
      for (const run of runs.items) {
        run.textRange.load("text");
      }
      await context.sync();
      for (const run of runs.items) {
        runTexts.push(run.textRange.text);
      }

      const paragraphText = para.textRange.text;
      const state = classifyBinding(binding, runTexts, paragraphText);
      classified.push({ binding, state });

      if (state === "broken") {
        brokenList.push({
          ...binding,
          slideNumber: binding.slideIndex + 1,
          shapeName,
        });
        continue;
      }

      // Apply update for clean or recoverable bindings
      if (state === "recoverable") {
        // Re-find the lastKnownValue in the paragraph and target it
        const offset = findLastKnownValueOffset(binding.lastKnownValue, paragraphText);
        if (offset === -1) {
          brokenList.push({ ...binding, slideNumber: binding.slideIndex + 1, shapeName });
          continue;
        }
        const targetRange = para.textRange.getSubstring(offset, binding.lastKnownValue.length);
        targetRange.text = variable.value;
        await context.sync();

        // Update binding record
        const bindingIdx = updatedBindings.findIndex((b) => b.id === binding.id);
        if (bindingIdx !== -1) {
          updatedBindings[bindingIdx] = {
            ...updatedBindings[bindingIdx],
            lastKnownValue: variable.value,
            charOffset: offset,
          };
        }
      } else {
        // Clean — update the specific run
        if (runs.items[binding.runIndex]) {
          runs.items[binding.runIndex].textRange.text = variable.value;
          await context.sync();
        } else {
          // Run index out of range, fall back to paragraph-level search
          const offset = findLastKnownValueOffset(binding.lastKnownValue, paragraphText);
          if (offset !== -1) {
            const targetRange = para.textRange.getSubstring(offset, binding.lastKnownValue.length);
            targetRange.text = variable.value;
            await context.sync();
          }
        }

        const bindingIdx = updatedBindings.findIndex((b) => b.id === binding.id);
        if (bindingIdx !== -1) {
          updatedBindings[bindingIdx] = {
            ...updatedBindings[bindingIdx],
            lastKnownValue: variable.value,
          };
        }
      }
    }
  });

  const clean = classified.filter((c) => c.state === "clean").length;
  const recovered = classified.filter((c) => c.state === "recoverable" &&
    !brokenList.find((b) => b.id === c.binding.id)).length;

  return {
    updatedBindings,
    summary: { clean, recovered, broken: brokenList },
  };
}

/**
 * Sync all variables in the registry.
 */
export async function syncAll(
  registry: VarSyncRegistry
): Promise<{ updatedRegistry: VarSyncRegistry; summary: SyncSummary }> {
  let currentBindings = [...registry.bindings];
  const aggregateSummary: SyncSummary = { clean: 0, recovered: 0, broken: [] };

  for (const variable of registry.variables) {
    const { updatedBindings, summary } = await syncVariable(variable, currentBindings);
    currentBindings = updatedBindings;
    aggregateSummary.clean += summary.clean;
    aggregateSummary.recovered += summary.recovered;
    aggregateSummary.broken.push(...summary.broken);
  }

  return {
    updatedRegistry: { ...registry, bindings: currentBindings },
    summary: aggregateSummary,
  };
}

/**
 * Re-learn a broken binding from the user's current selection.
 * Updates runIndex, charOffset, and lastKnownValue based on the new selection.
 */
export function reLearnBinding(
  binding: Binding,
  newRunIndex: number,
  newCharOffset: number,
  newValue: string
): Binding {
  return {
    ...binding,
    runIndex: newRunIndex,
    charOffset: newCharOffset,
    lastKnownValue: newValue,
  };
}
