import React, { useEffect, useMemo, useState } from "react";
import { VariablesPanel } from "./components/VariablesPanel";
import { LinkPanel } from "./components/LinkPanel";
import { BindingsPanel } from "./components/BindingsPanel";
import { HighlightToggle } from "./components/HighlightToggle";
import { SyncSummaryPanel } from "./components/SyncSummaryPanel";
import { useRegistry } from "./hooks/useRegistry";
import { syncAll } from "./lib/syncer";
import { useVarSyncStore } from "./store/useVarSyncStore";

type Tab = "variables" | "link" | "bindings" | "highlight";

export function App() {
  const [activeTab, setActiveTab] = useState<Tab>("variables");
  const [initialized, setInitialized] = useState(false);
  const [syncingAll, setSyncingAll] = useState(false);
  const { load, registry, updateAllBindings } = useRegistry();
  const { setSyncSummary } = useVarSyncStore();

  // Load registry from custom XML on mount
  useEffect(() => {
    load()
      .catch(console.error)
      .finally(() => setInitialized(true));
  }, []);

  // True when any binding is broken or out of date with its variable's current value.
  // Drives Sync All button prominence so it acts as an ambient "needs attention" indicator.
  const needsSync = useMemo(() => {
    return registry.bindings.some((b) => {
      if (!b.lastKnownValue) return true; // structurally broken
      const variable = registry.variables.find((v) => v.name === b.variableName);
      return variable ? b.lastKnownValue !== variable.value : false; // stale
    });
  }, [registry]);

  const handleSyncAll = async () => {
    setSyncingAll(true);
    setSyncSummary(null);
    try {
      const { updatedRegistry, summary } = await syncAll(registry);
      await updateAllBindings(updatedRegistry.bindings);
      setSyncSummary(summary);
    } catch (e) {
      console.error(e);
    } finally {
      setSyncingAll(false);
    }
  };

  const tabs: { id: Tab; label: string }[] = [
    { id: "variables", label: "Variables" },
    { id: "link", label: "Link" },
    { id: "bindings", label: "Bindings" },
    { id: "highlight", label: "Highlights" },
  ];

  if (!initialized) {
    return (
      <div className="min-h-screen bg-neutral-900 flex items-center justify-center">
        <p className="text-xs text-neutral-500">Loading VarSync…</p>
      </div>
    );
  }

  const noVariables = registry.variables.length === 0;

  return (
    <div className="min-h-screen bg-neutral-900 text-neutral-100 flex flex-col">
      {/* Header */}
      <div className="px-3 pt-3 pb-2 border-b border-neutral-800 flex items-center justify-between">
        <div>
          <h1 className="text-sm font-bold text-white tracking-tight">VarSync</h1>
          <p className="text-xs text-neutral-500">PowerPoint Variable Linking</p>
        </div>
        <button
          onClick={() => void handleSyncAll()}
          disabled={syncingAll || noVariables}
          title={needsSync ? "Bindings are out of date — click to sync" : "Sync all linked shapes"}
          className={`text-xs rounded px-2.5 py-1 transition-colors font-medium ${
            syncingAll || noVariables
              ? "bg-neutral-800 text-neutral-600"
              : needsSync
              ? "bg-cyan-500 hover:bg-cyan-400 text-white shadow-[0_0_8px_rgba(6,182,212,0.4)]"
              : "bg-cyan-700 hover:bg-cyan-600 text-white"
          }`}
        >
          {syncingAll ? "Syncing…" : needsSync ? "Sync All ●" : "Sync All"}
        </button>
      </div>

      {/* Sync Summary */}
      <div className="px-3 pt-2">
        <SyncSummaryPanel />
      </div>

      {/* Tab bar */}
      <div className="flex border-b border-neutral-800 px-1 mt-1">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            onClick={() => setActiveTab(tab.id)}
            className={`flex-1 py-2 text-xs font-medium transition-colors ${
              activeTab === tab.id
                ? "text-cyan-400 border-b-2 border-cyan-400"
                : "text-neutral-500 hover:text-neutral-300"
            }`}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* Tab content */}
      <div className="flex-1 overflow-y-auto px-3 py-3">
        {activeTab === "variables" && <VariablesPanel />}
        {activeTab === "link" && <LinkPanel />}
        {activeTab === "bindings" && <BindingsPanel />}
        {activeTab === "highlight" && <HighlightToggle />}
      </div>
    </div>
  );
}
