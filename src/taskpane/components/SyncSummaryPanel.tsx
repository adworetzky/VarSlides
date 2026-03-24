import React from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { BrokenLinkRow } from "./BrokenLinkRow";

export function SyncSummaryPanel() {
  const { syncSummary, setSyncSummary } = useVarSyncStore();

  if (!syncSummary) return null;

  const hasBroken = syncSummary.broken.length > 0;

  return (
    <div
      className={`border rounded p-2 flex flex-col gap-2 ${
        hasBroken
          ? "border-amber-700/50 bg-amber-950/20"
          : "border-green-800/50 bg-green-950/20"
      }`}
    >
      <div className="flex items-center justify-between">
        <span className="text-xs font-semibold text-neutral-300">Sync Result</span>
        <button
          onClick={() => setSyncSummary(null)}
          className="text-neutral-500 hover:text-neutral-300 text-xs transition-colors"
        >
          ×
        </button>
      </div>

      <div className="flex gap-3 text-xs">
        <span className="text-green-400">
          {syncSummary.clean} clean
        </span>
        {syncSummary.recovered > 0 && (
          <span className="text-cyan-400">
            {syncSummary.recovered} recovered
          </span>
        )}
        {hasBroken && (
          <span className="text-amber-400">
            {syncSummary.broken.length} broken
          </span>
        )}
      </div>

      {hasBroken && (
        <div className="flex flex-col gap-1.5">
          <p className="text-xs text-amber-300">
            The following links could not be updated automatically:
          </p>
          {syncSummary.broken.map((b) => (
            <BrokenLinkRow key={b.id} binding={b} />
          ))}
        </div>
      )}
    </div>
  );
}
