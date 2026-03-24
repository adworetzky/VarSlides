import React from "react";
import { useHighlight } from "../hooks/useHighlight";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { isColorCycling } from "../lib/highlighter";

export function HighlightToggle() {
  const { highlightMode, toggleHighlights } = useHighlight();
  const { registry } = useVarSyncStore();
  const cycling = isColorCycling(registry.variables);

  return (
    <div className="flex flex-col gap-3">
      {/* Toggle row */}
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
            Highlights
          </h2>
          <p className="text-xs text-neutral-500 mt-0.5">
            Color-code linked text in your slides for review
          </p>
        </div>
        <button
          onClick={() => void toggleHighlights()}
          className={`relative inline-flex h-5 w-9 items-center rounded-full transition-colors focus:outline-none flex-shrink-0 ${
            highlightMode ? "bg-cyan-600" : "bg-neutral-700"
          }`}
          role="switch"
          aria-checked={highlightMode}
          title={highlightMode ? "Turn off highlights" : "Turn on highlights"}
        >
          <span
            className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white shadow transition-transform ${
              highlightMode ? "translate-x-4.5" : "translate-x-0.5"
            }`}
          />
        </button>
      </div>

      {/* Active warning */}
      {highlightMode && (
        <div className="bg-amber-950/30 border border-amber-700/40 rounded px-2 py-1.5">
          <p className="text-xs text-amber-300 font-medium">Highlights are active</p>
          <p className="text-xs text-amber-400/80 mt-0.5">
            On Windows Desktop, they&apos;re stripped automatically on save.
            On Mac/Web, turn highlights off before saving.
          </p>
        </div>
      )}

      {/* Color legend */}
      {registry.variables.length > 0 ? (
        <div className="flex flex-col gap-1">
          <p className="text-xs text-neutral-500 uppercase tracking-wider font-semibold">
            Color legend
          </p>
          {registry.variables.map((v) => (
            <div key={v.name} className="flex items-center gap-2 text-xs">
              <span
                className="w-3 h-3 rounded-sm flex-shrink-0"
                style={{ backgroundColor: v.color }}
              />
              <span className="text-neutral-300 truncate">{v.name}</span>
              <span className="text-neutral-500 truncate ml-auto font-mono text-xs">
                {v.value}
              </span>
            </div>
          ))}
          {cycling && (
            <p className="text-xs text-neutral-500 mt-1">
              More than 8 variables — colors are repeating.
            </p>
          )}
        </div>
      ) : (
        <p className="text-xs text-neutral-500">
          Add variables to see the color legend.
        </p>
      )}
    </div>
  );
}
