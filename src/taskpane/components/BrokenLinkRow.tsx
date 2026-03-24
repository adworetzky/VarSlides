import React, { useState } from "react";
import type { BrokenBinding } from "../../types";
import { reLearnBinding } from "../lib/syncer";
import { useRegistry } from "../hooks/useRegistry";
import { useVarSyncStore } from "../store/useVarSyncStore";

interface BrokenLinkRowProps {
  binding: BrokenBinding;
}

export function BrokenLinkRow({ binding }: BrokenLinkRowProps) {
  const { removeBinding, updateBinding } = useRegistry();
  const { selectionContext } = useVarSyncStore();
  const [relearning, setRelearning] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleReLearn = async () => {
    if (!selectionContext) {
      setError("Select the correct text in the presentation first");
      return;
    }
    setRelearning(true);
    setError(null);
    try {
      // Capture current selection text as new lastKnownValue
      await PowerPoint.run(async (context) => {
        const selectedRange = context.presentation.getSelectedTextRange();
        selectedRange.load("text");
        await context.sync();

        const newValue = selectedRange.text;
        if (!newValue) throw new Error("No text selected");

        // Determine run index and char offset from selection
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        const slide = slides.items[selectionContext.slideIndex];
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        const shape = shapes.items.find((s) => s.id === selectionContext.shapeId);
        if (!shape) throw new Error("Shape not found");

        const paras = shape.textFrame.textRange.paragraphs;
        paras.load("items");
        await context.sync();

        let runIndex = 0;
        let charOffset = 0;
        for (const para of paras.items) {
          para.textRange.load("text");
          const runs = para.textRange.runs;
          runs.load("items");
          await context.sync();

          let cumChar = 0;
          for (let ri = 0; ri < runs.items.length; ri++) {
            runs.items[ri].textRange.load("text");
            await context.sync();
            const rt = runs.items[ri].textRange.text;
            if (rt === newValue) {
              runIndex = ri;
              charOffset = cumChar;
              break;
            }
            cumChar += rt.length;
          }
        }

        const updated = reLearnBinding(binding, runIndex, charOffset, newValue);
        await updateBinding(updated);
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setRelearning(false);
    }
  };

  const handleUnlink = async () => {
    await removeBinding(binding.id);
  };

  return (
    <div className="border border-amber-700/50 bg-amber-950/20 rounded p-2 flex flex-col gap-1">
      <div className="flex items-start justify-between gap-2">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1.5 text-xs">
            <span className="text-amber-400 font-medium">{binding.variableName}</span>
            <span className="text-neutral-500">·</span>
            <span className="text-neutral-400">Slide {binding.slideNumber}</span>
            <span className="text-neutral-500">·</span>
            <span className="text-neutral-400 truncate">{binding.shapeName}</span>
          </div>
          <p className="text-xs text-neutral-500 mt-0.5 truncate">
            Last known: <span className="font-mono text-amber-300">{binding.lastKnownValue}</span>
          </p>
        </div>
        <span className="text-xs text-amber-400 flex-shrink-0">Broken</span>
      </div>

      {error && (
        <p className="text-xs text-red-400">{error}</p>
      )}

      <div className="flex gap-1.5 mt-0.5">
        <button
          onClick={() => void handleReLearn()}
          disabled={relearning || !selectionContext}
          className="flex-1 bg-amber-700 hover:bg-amber-600 disabled:bg-neutral-800 disabled:text-neutral-600 text-white text-xs rounded py-1 transition-colors"
          title={!selectionContext ? "Select the correct text in PowerPoint first" : undefined}
        >
          {relearning ? "Re-learning…" : "Re-learn"}
        </button>
        <button
          onClick={() => void handleUnlink()}
          className="flex-1 bg-neutral-700 hover:bg-neutral-600 text-neutral-300 text-xs rounded py-1 transition-colors"
        >
          Unlink
        </button>
      </div>
    </div>
  );
}
