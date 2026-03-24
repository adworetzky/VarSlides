/**
 * highlighter.ts
 *
 * Apply and strip highlight formatting for VarSync's highlight mode.
 * Highlights must NEVER persist to saved files.
 */

import type { Binding, Variable, VarSyncRegistry } from "../../types";
import { HIGHLIGHT_PALETTE } from "../../types";

// ─────────────────────────────────────────────────────────────────────────────
// Color assignment (pure, unit-testable)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Assign a highlight color to each variable based on its position in the list.
 * Colors cycle if more than 8 variables exist.
 */
export function assignVariableColors(variables: Variable[]): Map<string, string> {
  const map = new Map<string, string>();
  variables.forEach((v, i) => {
    // Use the color already stored on the variable, or assign from palette
    map.set(v.name, v.color || HIGHLIGHT_PALETTE[i % HIGHLIGHT_PALETTE.length]);
  });
  return map;
}

/**
 * Returns true if color cycling is active (more than 8 variables).
 */
export function isColorCycling(variables: Variable[]): boolean {
  return variables.length > HIGHLIGHT_PALETTE.length;
}

/**
 * Get the highlight color for a variable by name from the registry.
 */
export function getVariableColor(variableName: string, variables: Variable[]): string {
  const v = variables.find((x) => x.name === variableName);
  if (v?.color) return v.color;
  const idx = variables.findIndex((x) => x.name === variableName);
  const safeIdx = idx < 0 ? 0 : idx;
  return HIGHLIGHT_PALETTE[safeIdx % HIGHLIGHT_PALETTE.length];
}

// ─────────────────────────────────────────────────────────────────────────────
// Hex color helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Convert #RRGGBB hex string to the integer format Office JS expects. */
function hexToOfficeColor(hex: string): string {
  // Office JS accepts hex strings directly for most APIs, return as-is
  return hex.replace("#", "");
}

// ─────────────────────────────────────────────────────────────────────────────
// Apply highlights (Office JS)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Apply highlight colors to all bound runs.
 * Stores the original colors in each binding record (mutates the registry copy
 * in memory; caller is responsible for persisting to custom XML).
 */
export async function applyHighlights(
  registry: VarSyncRegistry
): Promise<VarSyncRegistry> {
  const colorMap = assignVariableColors(registry.variables);
  const updatedBindings = [...registry.bindings];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (let i = 0; i < updatedBindings.length; i++) {
      const binding = updatedBindings[i];
      const color = colorMap.get(binding.variableName);
      if (!color) continue;

      const slide = slides.items[binding.slideIndex];
      if (!slide) continue;

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      const shape = shapes.items.find((s) => s.id === binding.shapeId);
      if (!shape) continue;

      const paras = shape.textFrame.textRange.paragraphs;
      paras.load("items");
      await context.sync();

      const para = paras.items[binding.paragraphIndex];
      if (!para) continue;

      const runs = para.textRange.runs;
      runs.load("items");
      await context.sync();

      const run = runs.items[binding.runIndex];
      if (!run) continue;

      // Load current colors before overwriting
      run.textRange.font.load("color, highlightColor");
      await context.sync();

      const origFontColor = (run.textRange.font.color as string) ?? null;
      const origHighlight = (run.textRange.font.highlightColor as string) ?? null;

      // Apply highlight
      (run.textRange.font as unknown as { highlightColor: string }).highlightColor =
        hexToOfficeColor(color);
      await context.sync();

      updatedBindings[i] = {
        ...binding,
        originalFontColor: origFontColor,
        originalHighlightColor: origHighlight,
      };
    }
  });

  return { ...registry, bindings: updatedBindings };
}

/**
 * Strip all highlights from bound runs, restoring original formatting.
 */
export async function stripHighlights(registry: VarSyncRegistry): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const binding of registry.bindings) {
      const slide = slides.items[binding.slideIndex];
      if (!slide) continue;

      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();

      const shape = shapes.items.find((s) => s.id === binding.shapeId);
      if (!shape) continue;

      const paras = shape.textFrame.textRange.paragraphs;
      paras.load("items");
      await context.sync();

      const para = paras.items[binding.paragraphIndex];
      if (!para) continue;

      const runs = para.textRange.runs;
      runs.load("items");
      await context.sync();

      const run = runs.items[binding.runIndex];
      if (!run) continue;

      // Restore original colors
      if (binding.originalFontColor !== null) {
        run.textRange.font.color = binding.originalFontColor;
      }
      const font = run.textRange.font as unknown as { highlightColor: string | null };
      font.highlightColor = binding.originalHighlightColor ?? null;
      await context.sync();
    }
  });
}
