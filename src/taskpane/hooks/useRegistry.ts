import { useCallback } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { loadRegistry, saveRegistry, cleanOrphanedBindings } from "../lib/registry";
import type { Variable, Binding, VarSyncRegistry, VariableSet } from "../../types";
import { HIGHLIGHT_PALETTE } from "../../types";

export function useRegistry() {
  const { registry, setRegistry } = useVarSyncStore();

  const load = useCallback(async () => {
    const loaded = await loadRegistry();
    setRegistry(loaded);
  }, [setRegistry]);

  // Optimistically updates the store; rolls back if the XML write fails.
  const persist = useCallback(
    async (next: VarSyncRegistry) => {
      const previous = registry;
      setRegistry(next);
      try {
        await saveRegistry(next);
      } catch (err) {
        setRegistry(previous);
        throw err;
      }
    },
    [registry, setRegistry]
  );

  const addVariable = useCallback(
    async (name: string, value: string) => {
      if (registry.variables.find((v) => v.name === name)) {
        throw new Error(`Variable "${name}" already exists`);
      }
      const colorIndex = registry.variables.length % HIGHLIGHT_PALETTE.length;
      const variable: Variable = {
        name,
        value,
        color: HIGHLIGHT_PALETTE[colorIndex],
        createdAt: new Date().toISOString(),
      };
      const next: VarSyncRegistry = {
        ...registry,
        variables: [...registry.variables, variable],
      };
      await persist(next);
    },
    [registry, persist]
  );

  const updateVariable = useCallback(
    async (name: string, value: string) => {
      const next: VarSyncRegistry = {
        ...registry,
        variables: registry.variables.map((v) =>
          v.name === name ? { ...v, value } : v
        ),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const deleteVariable = useCallback(
    async (name: string) => {
      const next: VarSyncRegistry = {
        ...registry,
        variables: registry.variables.filter((v) => v.name !== name),
        bindings: registry.bindings.filter((b) => b.variableName !== name),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const addBinding = useCallback(
    async (binding: Binding) => {
      const next: VarSyncRegistry = {
        ...registry,
        bindings: [...registry.bindings, binding],
      };
      await persist(next);
    },
    [registry, persist]
  );

  const removeBinding = useCallback(
    async (bindingId: string) => {
      const next: VarSyncRegistry = {
        ...registry,
        bindings: registry.bindings.filter((b) => b.id !== bindingId),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const updateBinding = useCallback(
    async (updated: Binding) => {
      const next: VarSyncRegistry = {
        ...registry,
        bindings: registry.bindings.map((b) => (b.id === updated.id ? updated : b)),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const updateAllBindings = useCallback(
    async (bindings: Binding[]) => {
      const next: VarSyncRegistry = { ...registry, bindings };
      await persist(next);
    },
    [registry, persist]
  );

  const renameVariable = useCallback(
    async (oldName: string, newName: string) => {
      if (registry.variables.find((v) => v.name === newName)) {
        throw new Error(`Variable "${newName}" already exists`);
      }
      const next: VarSyncRegistry = {
        ...registry,
        variables: registry.variables.map((v) =>
          v.name === oldName ? { ...v, name: newName } : v
        ),
        bindings: registry.bindings.map((b) =>
          b.variableName === oldName ? { ...b, variableName: newName } : b
        ),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const replaceRegistry = useCallback(
    async (next: VarSyncRegistry) => {
      await persist(next);
    },
    [persist]
  );

  const runCleanup = useCallback(async () => {
    const clean = cleanOrphanedBindings(registry);
    if (clean.bindings.length !== registry.bindings.length) {
      await persist(clean);
    }
  }, [registry, persist]);

  /** Snapshot all current variable values into a named set. Overwrites if name already exists. */
  const saveVariableSet = useCallback(
    async (setName: string) => {
      const values: Record<string, string> = {};
      for (const v of registry.variables) {
        values[v.name] = v.value;
      }
      const newSet: VariableSet = { name: setName, values, createdAt: new Date().toISOString() };
      const existing = registry.variableSets ?? [];
      const next: VarSyncRegistry = {
        ...registry,
        variableSets: existing.some((s) => s.name === setName)
          ? existing.map((s) => (s.name === setName ? newSet : s))
          : [...existing, newSet],
      };
      await persist(next);
    },
    [registry, persist]
  );

  /** Apply a saved set — updates matching variable values; variables not in the set are untouched. */
  const applyVariableSet = useCallback(
    async (setName: string) => {
      const set = (registry.variableSets ?? []).find((s) => s.name === setName);
      if (!set) throw new Error(`Variable set "${setName}" not found`);
      const next: VarSyncRegistry = {
        ...registry,
        variables: registry.variables.map((v) =>
          Object.prototype.hasOwnProperty.call(set.values, v.name)
            ? { ...v, value: set.values[v.name] }
            : v
        ),
      };
      await persist(next);
    },
    [registry, persist]
  );

  const deleteVariableSet = useCallback(
    async (setName: string) => {
      const next: VarSyncRegistry = {
        ...registry,
        variableSets: (registry.variableSets ?? []).filter((s) => s.name !== setName),
      };
      await persist(next);
    },
    [registry, persist]
  );

  return {
    registry,
    load,
    addVariable,
    updateVariable,
    deleteVariable,
    renameVariable,
    replaceRegistry,
    addBinding,
    removeBinding,
    updateBinding,
    updateAllBindings,
    runCleanup,
    saveVariableSet,
    applyVariableSet,
    deleteVariableSet,
  };
}
