import React, { useState } from "react";
import { useRegistry } from "../hooks/useRegistry";
import { BrokenLinkRow } from "./BrokenLinkRow";
import { classifyBinding } from "../lib/syncer";
import type { Binding, BrokenBinding } from "../../types";

export function BindingsPanel() {
  const { registry, removeBinding } = useRegistry();
  const [unlinking, setUnlinking] = useState<string | null>(null);

  const { bindings, variables } = registry;

  if (bindings.length === 0) {
    return (
      <div className="flex flex-col gap-2">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Bindings
        </h2>
        <p className="text-xs text-neutral-500 text-center py-4">
          No bindings yet. Link shapes or text selections in the Link panel.
        </p>
      </div>
    );
  }

  // Group bindings by variable name
  const grouped = new Map<string, Binding[]>();
  for (const b of bindings) {
    if (!grouped.has(b.variableName)) grouped.set(b.variableName, []);
    grouped.get(b.variableName)!.push(b);
  }

  const handleUnlink = async (id: string) => {
    setUnlinking(id);
    try {
      await removeBinding(id);
    } finally {
      setUnlinking(null);
    }
  };

  // Identify broken bindings (state determined from lastKnownValue heuristics without live Office context)
  // Show them at the top
  const brokenBindings: BrokenBinding[] = bindings
    .filter((b) => {
      // Without live paragraph text, we use a heuristic: flag as potentially broken
      // only when lastKnownValue is empty string (which is always broken)
      return !b.lastKnownValue;
    })
    .map((b) => ({
      ...b,
      slideNumber: b.slideIndex + 1,
      shapeName: b.shapeId,
    }));

  return (
    <div className="flex flex-col gap-2">
      <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
        Bindings
        <span className="ml-1.5 text-neutral-500 font-normal normal-case">
          ({bindings.length})
        </span>
      </h2>

      {/* Broken bindings at top */}
      {brokenBindings.length > 0 && (
        <div className="flex flex-col gap-1.5">
          {brokenBindings.map((b) => (
            <BrokenLinkRow key={b.id} binding={b} />
          ))}
        </div>
      )}

      {/* Grouped by variable */}
      {Array.from(grouped.entries()).map(([varName, varBindings]) => {
        const variable = variables.find((v) => v.name === varName);
        return (
          <div key={varName} className="flex flex-col gap-1">
            <div className="flex items-center gap-1.5">
              {variable && (
                <span
                  className="w-2 h-2 rounded-full flex-shrink-0"
                  style={{ backgroundColor: variable.color }}
                />
              )}
              <span className="text-xs font-medium text-neutral-300">{varName}</span>
              <span className="text-xs text-neutral-500">
                {varBindings.length} binding{varBindings.length !== 1 ? "s" : ""}
              </span>
            </div>

            {varBindings.map((b) => (
              <div
                key={b.id}
                className="ml-3 bg-neutral-800/40 border border-neutral-700/60 rounded p-1.5 flex items-start gap-2"
              >
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-1.5 text-xs">
                    <span className="text-neutral-400">Slide {b.slideIndex + 1}</span>
                    <span className="text-neutral-600">·</span>
                    <span className="text-neutral-400 truncate">{b.shapeId}</span>
                  </div>
                  <p className="text-xs font-mono text-neutral-300 truncate mt-0.5">
                    {b.lastKnownValue || (
                      <span className="text-red-400 italic">empty — broken</span>
                    )}
                  </p>
                  <p className="text-xs text-neutral-600 mt-0.5">
                    para {b.paragraphIndex}, run {b.runIndex}, char {b.charOffset}
                  </p>
                </div>
                <button
                  onClick={() => void handleUnlink(b.id)}
                  disabled={unlinking === b.id}
                  className="text-xs text-neutral-500 hover:text-red-400 transition-colors flex-shrink-0 pt-0.5"
                  title="Remove binding"
                >
                  {unlinking === b.id ? "…" : "Unlink"}
                </button>
              </div>
            ))}
          </div>
        );
      })}
    </div>
  );
}
