import React, { useState } from "react";
import { useSelection } from "../hooks/useSelection";
import { useRegistry } from "../hooks/useRegistry";
import { linkWholeShape, linkInlineSelection, isLinkableShapeType } from "../lib/linker";
import type { Variable } from "../../types";

export function LinkPanel() {
  const { selectionContext, refreshSelection } = useSelection();
  const { registry, addBinding } = useRegistry();

  const [selectedVarName, setSelectedVarName] = useState<string>("");
  const [error, setError] = useState<string | null>(null);
  const [linking, setLinking] = useState(false);
  const [success, setSuccess] = useState<string | null>(null);

  const selectedVariable: Variable | undefined = registry.variables.find(
    (v) => v.name === selectedVarName
  );

  const handleLinkShape = async () => {
    if (!selectionContext || !selectedVariable) return;
    setError(null);
    setSuccess(null);
    setLinking(true);
    try {
      const binding = await linkWholeShape(
        selectedVariable,
        selectionContext.shapeId,
        selectionContext.shapeName,
        selectionContext.slideIndex
      );
      await addBinding(binding);
      setSuccess(`Linked shape "${selectionContext.shapeName}" to ${selectedVariable.name}`);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLinking(false);
    }
  };

  const handleLinkSelection = async () => {
    if (!selectionContext || !selectedVariable || !selectionContext.hasTextSelection) return;
    setError(null);
    setSuccess(null);
    setLinking(true);
    try {
      // Get the selection range from the Office JS context
      await PowerPoint.run(async (context) => {
        const selectedRange = context.presentation.getSelectedTextRange();
        selectedRange.load("text");
        await context.sync();

        // We need paragraph index and character offsets.
        // These require traversing the shape's paragraph structure.
        // For now, link at paragraph 0 with best-effort offsets.
        // Full implementation requires Phase 3 Script Lab validation.
        const selectedText = selectedRange.text;
        if (!selectedText) throw new Error("No text selected");

        const shapes = context.presentation.slides.items;
        // Find the paragraph index and char offset
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[selectionContext.slideIndex];
        const slideShapes = slide.shapes;
        slideShapes.load("items");
        await context.sync();

        const shape = slideShapes.items.find((s) => s.id === selectionContext.shapeId);
        if (!shape) throw new Error("Shape not found");

        const paras = shape.textFrame.textRange.paragraphs;
        paras.load("items");
        await context.sync();

        let paragraphIndex = 0;
        let selectionStart = 0;
        for (let pi = 0; pi < paras.items.length; pi++) {
          paras.items[pi].textRange.load("text");
          await context.sync();
          const paraText = paras.items[pi].textRange.text;
          const idx = paraText.indexOf(selectedText);
          if (idx !== -1) {
            paragraphIndex = pi;
            selectionStart = idx;
            break;
          }
        }

        const binding = await linkInlineSelection(
          selectedVariable,
          selectionContext.shapeId,
          selectionContext.slideIndex,
          paragraphIndex,
          selectionStart,
          selectionStart + selectedText.length
        );
        await addBinding(binding);
        setSuccess(`Linked selection "${selectedText}" to ${selectedVariable.name}`);
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLinking(false);
    }
  };

  // Empty state
  if (!selectionContext) {
    return (
      <div className="flex flex-col gap-2">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Link
        </h2>
        <div className="border border-dashed border-neutral-700 rounded p-4 text-center">
          <p className="text-xs text-neutral-500">
            Select a shape or text in your presentation to link it to a variable.
          </p>
          <button
            onClick={() => void refreshSelection()}
            className="mt-2 text-xs text-cyan-500 hover:text-cyan-300 transition-colors"
          >
            Refresh selection
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-2">
      <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
        Link
      </h2>

      {/* Selection context */}
      <div className="bg-neutral-800/50 border border-neutral-700/80 rounded p-2">
        <div className="flex items-center gap-1.5 text-xs">
          <span className="text-neutral-400">Shape:</span>
          <span className="text-neutral-200 font-medium truncate">
            {selectionContext.shapeName}
          </span>
        </div>
        <div className="flex items-center gap-1.5 text-xs mt-0.5">
          <span className="text-neutral-400">Slide:</span>
          <span className="text-neutral-300">{selectionContext.slideIndex + 1}</span>
          {selectionContext.hasTextSelection && (
            <span className="ml-auto text-cyan-400 text-xs">Text selected</span>
          )}
        </div>
      </div>

      {/* Variable selector */}
      {registry.variables.length === 0 ? (
        <p className="text-xs text-neutral-500">
          No variables defined. Add one in the Variables panel.
        </p>
      ) : (
        <select
          className="bg-neutral-800 border border-neutral-700 rounded px-2 py-1.5 text-xs text-neutral-200 focus:outline-none focus:border-cyan-500 w-full"
          value={selectedVarName}
          onChange={(e) => setSelectedVarName(e.target.value)}
        >
          <option value="">Select a variable…</option>
          {registry.variables.map((v) => (
            <option key={v.name} value={v.name}>
              {v.name} = {v.value}
            </option>
          ))}
        </select>
      )}

      {/* Action buttons */}
      <div className="flex gap-1.5">
        <button
          onClick={() => void handleLinkShape()}
          disabled={!selectedVariable || linking}
          className="flex-1 bg-neutral-700 hover:bg-neutral-600 disabled:bg-neutral-800 disabled:text-neutral-600 text-neutral-200 text-xs rounded py-1.5 transition-colors font-medium"
        >
          {linking ? "Linking…" : "Link Shape"}
        </button>
        <button
          onClick={() => void handleLinkSelection()}
          disabled={!selectedVariable || !selectionContext.hasTextSelection || linking}
          className="flex-1 bg-cyan-700 hover:bg-cyan-600 disabled:bg-neutral-800 disabled:text-neutral-600 text-white text-xs rounded py-1.5 transition-colors font-medium"
          title={
            !selectionContext.hasTextSelection
              ? "Select text in the shape first"
              : undefined
          }
        >
          {linking ? "Linking…" : "Link Selection"}
        </button>
      </div>

      {!selectionContext.hasTextSelection && selectedVariable && (
        <p className="text-xs text-neutral-500">
          To link a substring, enter text editing mode in PowerPoint and select the
          specific text, then click Link Selection.
        </p>
      )}

      {/* Note about mixed formatting */}
      {selectionContext.hasTextSelection && (
        <p className="text-xs text-amber-500/80">
          Note: The linked run inherits formatting from the original run. If the
          selection had different formatting, set it manually after linking.
        </p>
      )}

      {/* Error / Success feedback */}
      {error && (
        <div className="text-xs text-red-400 bg-red-950/40 border border-red-800/50 rounded px-2 py-1">
          {error}
        </div>
      )}
      {success && (
        <div className="text-xs text-green-400 bg-green-950/40 border border-green-800/50 rounded px-2 py-1">
          {success}
        </div>
      )}
    </div>
  );
}
