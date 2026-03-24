import React, { useState, useRef } from "react";
import { useRegistry } from "../hooks/useRegistry";
import { syncVariable } from "../lib/syncer";
import { useVarSyncStore } from "../store/useVarSyncStore";
import type { Variable } from "../../types";

interface EditState {
  name: string;
  value: string;
}

export function VariablesPanel() {
  const { registry, addVariable, updateVariable, deleteVariable, updateAllBindings } =
    useRegistry();
  const { setSyncSummary } = useVarSyncStore();

  const [editingName, setEditingName] = useState<string | null>(null);
  const [editValue, setEditValue] = useState("");
  const [newVarName, setNewVarName] = useState("");
  const [newVarValue, setNewVarValue] = useState("");
  const [showAddForm, setShowAddForm] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [syncing, setSyncing] = useState<string | null>(null);
  const [deletingName, setDeletingName] = useState<string | null>(null);
  const editInputRef = useRef<HTMLInputElement>(null);

  const bindingCountFor = (name: string) =>
    registry.bindings.filter((b) => b.variableName === name).length;

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

  return (
    <div className="flex flex-col gap-2">
      {/* Header */}
      <div className="flex items-center justify-between">
        <h2 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">
          Variables
        </h2>
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

      {/* Error */}
      {error && (
        <div className="text-xs text-red-400 bg-red-950/40 border border-red-800/50 rounded px-2 py-1">
          {error}
          <button
            onClick={() => setError(null)}
            className="ml-2 text-red-300 hover:text-red-100"
          >
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

      {/* Variable List */}
      {registry.variables.length === 0 && !showAddForm && (
        <p className="text-xs text-neutral-500 text-center py-4">
          No variables yet. Click + Add to create one.
        </p>
      )}

      {registry.variables.map((v) => {
        const count = bindingCountFor(v.name);
        const isEditing = editingName === v.name;
        const isDeleting = deletingName === v.name;

        return (
          <div
            key={v.name}
            className="bg-neutral-800/50 border border-neutral-700/80 rounded p-2 flex flex-col gap-1"
          >
            <div className="flex items-center gap-2">
              {/* Color chip */}
              <span
                className="w-2.5 h-2.5 rounded-full flex-shrink-0"
                style={{ backgroundColor: v.color }}
              />
              {/* Name */}
              <span className="text-xs font-medium text-neutral-200 flex-1 truncate">
                {v.name}
              </span>
              {/* Binding count */}
              <span className="text-xs text-neutral-500">{count} link{count !== 1 ? "s" : ""}</span>
              {/* Sync button */}
              <button
                onClick={() => void handleSync(v)}
                disabled={syncing !== null}
                className="text-xs text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 transition-colors"
                title="Sync this variable"
              >
                {syncing === v.name ? "..." : "Sync"}
              </button>
              {/* Delete button */}
              <button
                onClick={() => handleDeleteRequest(v.name)}
                className="text-xs text-neutral-500 hover:text-red-400 transition-colors"
                title="Delete variable"
              >
                ×
              </button>
            </div>

            {/* Value — inline edit */}
            {isEditing ? (
              <input
                ref={editInputRef}
                className="bg-neutral-900 border border-cyan-600 rounded px-2 py-0.5 text-xs text-neutral-100 focus:outline-none w-full"
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
                className="text-left text-xs text-cyan-300 font-mono hover:text-cyan-200 transition-colors truncate"
                onClick={() => handleStartEdit(v)}
                title="Click to edit value"
              >
                {v.value || <span className="text-neutral-500 italic">empty</span>}
              </button>
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
