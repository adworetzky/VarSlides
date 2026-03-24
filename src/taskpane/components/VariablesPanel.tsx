import React, { useMemo, useState, useRef } from "react";
import { useRegistry } from "../hooks/useRegistry";
import { syncVariable } from "../lib/syncer";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { FindLinkPanel } from "./FindLinkPanel";
import type { Variable, VarSyncRegistry } from "../../types";

type VariableHealth = "ok" | "broken" | "unknown";

export function VariablesPanel() {
  const {
    registry,
    addVariable,
    updateVariable,
    deleteVariable,
    renameVariable,
    replaceRegistry,
    updateAllBindings,
  } = useRegistry();
  const { setSyncSummary } = useVarSyncStore();

  const [filterText, setFilterText] = useState("");
  const [editingName, setEditingName] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [renamingName, setRenamingName] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [newVarName, setNewVarName] = useState("");
  const [newVarValue, setNewVarValue] = useState("");
  const [showAddForm, setShowAddForm] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [syncing, setSyncing] = useState<string | null>(null);
  const [deletingName, setDeletingName] = useState<string | null>(null);
  const [findLinkVar, setFindLinkVar] = useState<Variable | null>(null);
  const editInputRef = useRef<HTMLInputElement>(null);
  const renameInputRef = useRef<HTMLInputElement>(null);
  const importInputRef = useRef<HTMLInputElement>(null);

  const bindingCounts = useMemo(() => {
    const counts = new Map<string, number>();
    for (const b of registry.bindings) {
      counts.set(b.variableName, (counts.get(b.variableName) ?? 0) + 1);
    }
    return counts;
  }, [registry.bindings]);

  const variableHealth = useMemo(() => {
    const health = new Map<string, VariableHealth>();
    for (const v of registry.variables) {
      const vb = registry.bindings.filter((b) => b.variableName === v.name);
      if (vb.length === 0) health.set(v.name, "unknown");
      else if (vb.some((b) => !b.lastKnownValue)) health.set(v.name, "broken");
      else health.set(v.name, "ok");
    }
    return health;
  }, [registry.variables, registry.bindings]);

  const filteredVariables = useMemo(() => {
    if (!filterText) return registry.variables;
    const lower = filterText.toLowerCase();
    return registry.variables.filter(
      (v) => v.name.toLowerCase().includes(lower) || v.value.toLowerCase().includes(lower)
    );
  }, [registry.variables, filterText]);

  const handleStartEdit = (v: Variable) => {
    setEditingName(v.name);
    setEditValue(v.value);
    setTimeout(() => editInputRef.current?.focus(), 0);
  };

  const handleCommitEdit = async (name: string) => {
    try {
      await updateVariable(name, editValue);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setEditingName(null);
    }
  };

  const handleStartRename = (v: Variable) => {
    setRenamingName(v.name);
    setRenameValue(v.name);
    setTimeout(() => renameInputRef.current?.focus(), 0);
  };

  const handleCommitRename = async (oldName: string) => {
    const trimmed = renameValue.trim();
    if (trimmed && trimmed !== oldName) {
      try {
        await renameVariable(oldName, trimmed);
        if (findLinkVar?.name === oldName) setFindLinkVar(null);
      } catch (e) {
        setError((e as Error).message);
      }
    }
    setRenamingName(null);
  };

  const handleAddVariable = async () => {
    const trimmedName = newVarName.trim();
    const trimmedValue = newVarValue.trim();
    if (!trimmedName) {
      setError("Variable name is required");
      return;
    }
    try {
      await addVariable(trimmedName, trimmedValue);
      setNewVarName("");
      setNewVarValue("");
      setShowAddForm(false);
      setError(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const handleDeleteRequest = (name: string) => {
    setDeletingName(name);
  };

  const handleConfirmDelete = async (name: string) => {
    try {
      await deleteVariable(name);
      if (findLinkVar?.name === name) setFindLinkVar(null);
      setDeletingName(null);
    } catch (e) {
      setError((e as Error).message);
      setDeletingName(null);
    }
  };

  const handleSync = async (variable: Variable) => {
    setSyncing(variable.name);
    setSyncSummary(null);
    try {
      const { updatedBindings, summary } = await syncVariable(variable, registry.bindings);
      await updateAllBindings(updatedBindings);
      setSyncSummary(summary);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setSyncing(null);
    }
  };

  const handleExport = () => {
    const json = JSON.stringify(registry, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "varsync-registry.json";
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as VarSyncRegistry;
      if (!Array.isArray(parsed.variables) || !Array.isArray(parsed.bindings)) {
        throw new Error("Invalid registry format");
      }
      await replaceRegistry(parsed);
      setFindLinkVar(null);
      setError(null);
    } catch (e) {
      setError(`Import failed: ${(e as Error).message}`);
    } finally {
      if (importInputRef.current) importInputRef.current.value = "";
    }
  };

  const healthDotClass: Record<VariableHealth, string> = {
    ok: "text-green-400",
    broken: "text-amber-400",
    unknown: "text-neutral-600",
  };

  const healthTitle: Record<VariableHealth, string> = {
    ok: "All bindings healthy",
    broken: "Some bindings broken — run Sync",
    unknown: "No bindings yet",
  };

  return (
    <div className="flex flex-col gap-2">
      {/* Header */}
      <div className="flex items-center justify-between">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Variables
        </h2>
        <div className="flex items-center gap-2">
          <button
            onClick={handleExport}
            disabled={registry.variables.length === 0}
            className="text-xs text-neutral-500 hover:text-neutral-300 disabled:text-neutral-700 transition-colors"
            title="Export registry as JSON"
          >
            Export
          </button>
          <label
            className="text-xs text-neutral-500 hover:text-neutral-300 transition-colors cursor-pointer"
            title="Import registry from JSON"
          >
            Import
            <input
              ref={importInputRef}
              type="file"
              accept=".json"
              className="hidden"
              onChange={(e) => void handleImport(e)}
            />
          </label>
          <button
            onClick={() => {
              setShowAddForm(true);
              setError(null);
            }}
            className="text-xs text-cyan-400 hover:text-cyan-300 transition-colors font-medium"
          >
            + Add
          </button>
        </div>
      </div>

      {/* Error */}
      {error && (
        <div className="text-xs text-red-400 bg-red-950/40 border border-red-800/50 rounded px-2 py-1">
          {error}
          <button onClick={() => setError(null)} className="ml-2 text-red-300 hover:text-red-100">
            ×
          </button>
        </div>
      )}

      {/* Add Form */}
      {showAddForm && (
        <div className="bg-neutral-800/60 border border-neutral-700 rounded p-2 flex flex-col gap-1.5">
          <input
            className="bg-neutral-900 border border-neutral-700 rounded px-2 py-1 text-xs text-neutral-100 placeholder-neutral-500 focus:outline-none focus:border-cyan-500"
            placeholder="Variable name (e.g. AUM)"
            value={newVarName}
            onChange={(e) => setNewVarName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && void handleAddVariable()}
          />
          <input
            className="bg-neutral-900 border border-neutral-700 rounded px-2 py-1 text-xs text-neutral-100 placeholder-neutral-500 focus:outline-none focus:border-cyan-500"
            placeholder="Value (e.g. $70B)"
            value={newVarValue}
            onChange={(e) => setNewVarValue(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && void handleAddVariable()}
          />
          <div className="flex gap-1.5">
            <button
              onClick={() => void handleAddVariable()}
              className="flex-1 bg-cyan-600 hover:bg-cyan-500 text-white text-xs rounded py-1 transition-colors font-medium"
            >
              Add Variable
            </button>
            <button
              onClick={() => {
                setShowAddForm(false);
                setNewVarName("");
                setNewVarValue("");
              }}
              className="px-2 text-xs text-neutral-400 hover:text-neutral-200 transition-colors"
            >
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* Search filter — only shown when there are enough variables to warrant it */}
      {registry.variables.length > 3 && (
        <input
          className="bg-neutral-900 border border-neutral-700 rounded px-2 py-1 text-xs text-neutral-200 placeholder-neutral-500 focus:outline-none focus:border-cyan-500"
          placeholder="Filter variables…"
          value={filterText}
          onChange={(e) => setFilterText(e.target.value)}
        />
      )}

      {/* Empty state */}
      {registry.variables.length === 0 && !showAddForm && (
        <p className="text-xs text-neutral-500 text-center py-4">
          No variables yet. Click + Add to create one.
        </p>
      )}

      {filteredVariables.length === 0 && filterText && (
        <p className="text-xs text-neutral-500 text-center py-2">
          No variables match &ldquo;{filterText}&rdquo;
        </p>
      )}

      {/* Variable List */}
      {filteredVariables.map((v) => {
        const count = bindingCounts.get(v.name) ?? 0;
        const isEditing = editingName === v.name;
        const isRenaming = renamingName === v.name;
        const isDeleting = deletingName === v.name;
        const health = variableHealth.get(v.name) ?? "unknown";
        const showFindLink = findLinkVar?.name === v.name;

        return (
          <div
            key={v.name}
            className="bg-neutral-800/50 border border-neutral-700/80 rounded p-2 flex flex-col gap-1"
          >
            {/* Name row */}
            {isRenaming ? (
              <div className="flex items-center gap-1.5">
                <span
                  className="w-2.5 h-2.5 rounded-full flex-shrink-0"
                  style={{ backgroundColor: v.color }}
                />
                <input
                  ref={renameInputRef}
                  className="flex-1 bg-neutral-900 border border-cyan-600 rounded px-2 py-0.5 text-xs text-neutral-100 focus:outline-none"
                  value={renameValue}
                  onChange={(e) => setRenameValue(e.target.value)}
                  onBlur={() => void handleCommitRename(v.name)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") void handleCommitRename(v.name);
                    if (e.key === "Escape") setRenamingName(null);
                  }}
                />
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <span
                  className="w-2.5 h-2.5 rounded-full flex-shrink-0"
                  style={{ backgroundColor: v.color }}
                />
                <button
                  className="text-xs font-medium text-neutral-200 hover:text-neutral-100 text-left truncate flex-1"
                  onClick={() => handleStartRename(v)}
                  title="Click to rename"
                >
                  {v.name}
                </button>
                <span
                  className={`text-xs leading-none ${healthDotClass[health]}`}
                  title={healthTitle[health]}
                >
                  {health === "unknown" ? "○" : "●"}
                </span>
                <span className="text-xs text-neutral-500">
                  {count} link{count !== 1 ? "s" : ""}
                </span>
                <button
                  onClick={() => void handleSync(v)}
                  disabled={syncing !== null}
                  className="text-xs text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 transition-colors"
                  title="Sync this variable"
                >
                  {syncing === v.name ? "..." : "Sync"}
                </button>
                <button
                  onClick={() => handleDeleteRequest(v.name)}
                  className="text-xs text-neutral-500 hover:text-red-400 transition-colors"
                  title="Delete variable"
                >
                  ×
                </button>
              </div>
            )}

            {/* Value row — inline edit + Find & Link toggle */}
            <div className="flex items-center gap-1.5">
              {isEditing ? (
                <input
                  ref={editInputRef}
                  className="flex-1 bg-neutral-900 border border-cyan-600 rounded px-2 py-0.5 text-xs text-neutral-100 focus:outline-none"
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  onBlur={() => void handleCommitEdit(v.name)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") void handleCommitEdit(v.name);
                    if (e.key === "Escape") setEditingName(null);
                  }}
                />
              ) : (
                <button
                  className="flex-1 text-left text-xs text-cyan-300 font-mono hover:text-cyan-200 transition-colors truncate"
                  onClick={() => handleStartEdit(v)}
                  title="Click to edit value"
                >
                  {v.value || <span className="text-neutral-500 italic">empty</span>}
                </button>
              )}
              <button
                onClick={() => setFindLinkVar(showFindLink ? null : v)}
                className={`text-xs flex-shrink-0 transition-colors ${
                  showFindLink
                    ? "text-cyan-300 font-medium"
                    : "text-neutral-500 hover:text-cyan-400"
                }`}
                title="Find & Link occurrences in the deck"
              >
                Find
              </button>
            </div>

            {/* Find & Link inline panel */}
            {showFindLink && (
              <FindLinkPanel variable={v} onClose={() => setFindLinkVar(null)} />
            )}

            {/* Delete confirmation */}
            {isDeleting && (
              <div className="mt-1 bg-red-950/40 border border-red-800/50 rounded px-2 py-1.5 flex items-center justify-between gap-2">
                <span className="text-xs text-red-300">
                  Delete variable and {count} binding{count !== 1 ? "s" : ""}?
                </span>
                <div className="flex gap-1.5">
                  <button
                    onClick={() => void handleConfirmDelete(v.name)}
                    className="text-xs bg-red-700 hover:bg-red-600 text-white rounded px-2 py-0.5 transition-colors"
                  >
                    Delete
                  </button>
                  <button
                    onClick={() => setDeletingName(null)}
                    className="text-xs text-neutral-400 hover:text-neutral-200 transition-colors"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
