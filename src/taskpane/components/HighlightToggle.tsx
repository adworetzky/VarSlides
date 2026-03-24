import React from "react";
import { useHighlight } from "../hooks/useHighlight";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { isColorCycling } from "../lib/highlighter";

export function HighlightToggle() {
  const { highlightMode, toggleHighlights } = useHighlight();
  const { registry } = useVarSyncStore();
  const cycling = isColorCycling(registry.variables);

  return (
    <div className="flex flex-col gap-2">
      <div className="flex items-center justify-between">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Highlights
        </h2>

        {/* Toggle */}
        <button
          onClick={() => void toggleHighlights()}
          className={`relative inline-flex h-5 w-9 items-center rounded-full transition-colors focus:outline-none ${
            highlightMode ? "bg-cyan-600" : "bg-neutral-700"
          }`}
          role="switch"
          aria-checked={highlightMode}
        >
          <span
            className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white shadow transition-transform ${
              highlightMode ? "translate-x-4.5" : "translate-x-0.5"
            }`}
          />
        </button>
      </div>

      {/* Save warning */}
      {highlightMode && (
        <div className="bg-amber-950/30 border border-amber-700/40 rounded px-2 py-1.5">
          <p className="text-xs text-amber-300">
            Highlights are active. Disable before saving to prevent them from
            appearing in the exported file.
          </p>
        </div>
      )}

      {/* Color legend */}
      {registry.variables.length > 0 && (
        <div className="flex flex-col gap-1">
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
              More than 8 variables — colors are cycling.
            </p>
          )}
        </div>
      )}

      {registry.variables.length === 0 && (
        <p className="text-xs text-neutral-500">
          Add variables to see the color legend.
        </p>
      )}
    </div>
  );
}
