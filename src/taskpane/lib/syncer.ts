import type {
  Binding,
  BrokenBinding,
  ClassifiedBinding,
  BindingState,
  VarSyncRegistry,
  SyncSummary,
  Variable,
} from "../../types";

export function classifyBinding(
  binding: Binding,
  runTexts: string[],
  paragraphText: string
): BindingState {
  const lkv = binding.lastKnownValue;
  if (runTexts.length > binding.runIndex && runTexts[binding.runIndex] === lkv) {
    return "clean";
  }
  if (paragraphText.includes(lkv)) {
    return "recoverable";
  }
  return "broken";
}

export function findLastKnownValueOffset(
  lastKnownValue: string,
  paragraphText: string
): number {
  return paragraphText.indexOf(lastKnownValue);
}

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

      for (const run of runs.items) {
        run.textRange.load("text");
      }
      await context.sync();

      const runTexts = runs.items.map((r) => r.textRange.text);

      const paragraphText = para.textRange.text;
      const state = classifyBinding(binding, runTexts, paragraphText);
      classified.push({ binding, state });

      if (state === "broken") {
        brokenList.push({
          ...binding,
          slideNumber: binding.slideIndex + 1,
          shapeName,
          currentText: paragraphText,
        });
        continue;
      }

      if (state === "recoverable") {
        const offset = findLastKnownValueOffset(binding.lastKnownValue, paragraphText);
        if (offset === -1) {
          brokenList.push({
            ...binding,
            slideNumber: binding.slideIndex + 1,
            shapeName,
            currentText: paragraphText,
          });
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
        if (runs.items[binding.runIndex]) {
          runs.items[binding.runIndex].textRange.text = variable.value;
          await context.sync();
        } else {
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
