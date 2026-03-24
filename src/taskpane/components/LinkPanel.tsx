import React, { useState } from "react";
import { useSelection } from "../hooks/useSelection";
import { useRegistry } from "../hooks/useRegistry";
import { linkWholeShape, linkInlineSelection } from "../lib/linker";
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

  const alreadyLinked = !!(
    selectedVariable &&
    selectionContext &&
    registry.bindings.some(
      (b) => b.variableName === selectedVarName && b.shapeId === selectionContext.shapeId
    )
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
      setSuccess(`Linked "${selectionContext.shapeName}" → ${selectedVariable.name}`);
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
      await PowerPoint.run(async (context) => {
        const selectedRange = context.presentation.getSelectedTextRange();
        selectedRange.load("text");
        await context.sync();

        const selectedText = selectedRange.text;
        if (!selectedText) throw new Error("No text selected");

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

        // Batch-load all paragraph texts before searching
        for (const para of paras.items) {
          para.textRange.load("text");
        }
        await context.sync();

        let paragraphIndex = 0;
        let selectionStart = 0;
        for (let pi = 0; pi < paras.items.length; pi++) {
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
          selectionStart + selectedText.length,
          selectionContext.shapeName
        );
        await addBinding(binding);
        setSuccess(`Linked "${selectedText}" → ${selectedVariable.name}`);
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLinking(false);
    }
  };

  // Empty state — selection auto-refreshes via DocumentSelectionChanged event
  if (!selectionContext) {
    return (
      <div className="flex flex-col gap-2">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Link
        </h2>
        <div className="border border-dashed border-neutral-700 rounded p-4 text-center flex flex-col gap-2">
          <p className="text-xs text-neutral-500">
            Click a shape in your presentation — this panel updates automatically.
          </p>
          <button
            onClick={() => void refreshSelection()}
            className="text-xs text-cyan-500 hover:text-cyan-300 transition-colors"
          >
            Re-check selection
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
          <span className="text-neutral-200 font-medium truncate flex-1">
            {selectionContext.shapeName}
          </span>
          <span className="text-neutral-500">Slide {selectionContext.slideIndex + 1}</span>
        </div>
        {selectionContext.hasTextSelection ? (
          <p className="text-xs text-cyan-400 mt-0.5">Text selected — ready to link a substring</p>
        ) : (
          <p className="text-xs text-neutral-500 mt-0.5">
            No text selected — Link Shape will replace the whole shape&apos;s content.
            To link a substring, double-click into the shape, select the text, then return here.
          </p>
        )}
      </div>

      {/* Variable selector */}
      {registry.variables.length === 0 ? (
        <p className="text-xs text-neutral-500">
          No variables yet — add one in the Variables tab first.
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

      {alreadyLinked && (
        <p className="text-xs text-amber-400 bg-amber-950/30 border border-amber-800/40 rounded px-2 py-1">
          This shape already has a binding for <span className="font-medium">{selectedVarName}</span>. Linking again will create a second binding.
        </p>
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
        >
          {linking ? "Linking…" : "Link Selection"}
        </button>
      </div>

      {/* Formatting note for inline links */}
      {selectionContext.hasTextSelection && selectedVariable && (
        <p className="text-xs text-neutral-500">
          The linked run inherits the original text&apos;s formatting. Adjust manually in PowerPoint if needed.
        </p>
      )}

      {/* Error / Success feedback */}
      {error && (
        <div className="text-xs text-red-400 bg-red-950/40 border border-red-800/50 rounded px-2 py-1 flex items-start gap-1.5">
          <span className="flex-1">{error}</span>
          <button onClick={() => setError(null)} className="text-red-300 hover:text-red-100 flex-shrink-0">×</button>
        </div>
      )}
      {success && (
        <div className="text-xs text-green-400 bg-green-950/40 border border-green-800/50 rounded px-2 py-1 flex items-start gap-1.5">
          <span className="flex-1">{success}</span>
          <button onClick={() => setSuccess(null)} className="text-green-300 hover:text-green-100 flex-shrink-0">×</button>
        </div>
      )}
    </div>
  );
}
