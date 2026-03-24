import React, { useEffect, useState } from "react";
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
          disabled={syncingAll || registry.variables.length === 0}
          className="text-xs bg-cyan-700 hover:bg-cyan-600 disabled:bg-neutral-800 disabled:text-neutral-600 text-white rounded px-2.5 py-1 transition-colors font-medium"
        >
          {syncingAll ? "Syncing…" : "Sync All"}
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
