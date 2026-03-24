import React, { useEffect } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { BrokenLinkRow } from "./BrokenLinkRow";

export function SyncSummaryPanel() {
  const { syncSummary, setSyncSummary } = useVarSyncStore();

  // Auto-dismiss after 3 s when everything is clean
  useEffect(() => {
    if (!syncSummary || syncSummary.broken.length > 0) return;
    const timer = setTimeout(() => setSyncSummary(null), 3000);
    return () => clearTimeout(timer);
  }, [syncSummary, setSyncSummary]);

  if (!syncSummary) return null;

  const hasBroken = syncSummary.broken.length > 0;
  const total = syncSummary.clean + syncSummary.recovered + syncSummary.broken.length;

  return (
    <div
      className={`border rounded p-2 flex flex-col gap-2 ${
        hasBroken
          ? "border-amber-700/50 bg-amber-950/20"
          : "border-green-800/50 bg-green-950/20"
      }`}
    >
      <div className="flex items-center justify-between">
        <span className="text-xs font-semibold text-neutral-300">
          {hasBroken ? "Sync complete — action needed" : "Sync complete"}
        </span>
        <button
          onClick={() => setSyncSummary(null)}
          className="text-neutral-500 hover:text-neutral-300 text-xs transition-colors"
        >
          ×
        </button>
      </div>

      <div className="flex gap-3 text-xs">
        <span className="text-green-400" title={`${syncSummary.clean} binding${syncSummary.clean !== 1 ? "s" : ""} updated normally`}>
          {syncSummary.clean}/{total} updated
        </span>
        {syncSummary.recovered > 0 && (
          <span className="text-cyan-400" title="Run was remerged; binding auto-repaired">
            {syncSummary.recovered} auto-repaired
          </span>
        )}
        {hasBroken && (
          <span className="text-amber-400" title="These bindings need manual attention">
            {syncSummary.broken.length} need attention
          </span>
        )}
      </div>

      {hasBroken && (
        <div className="flex flex-col gap-1.5">
          <p className="text-xs text-amber-300">
            These links could not be updated — the text has changed or been deleted:
          </p>
          {syncSummary.broken.map((b) => (
            <BrokenLinkRow key={b.id} binding={b} />
          ))}
        </div>
      )}
    </div>
  );
}
