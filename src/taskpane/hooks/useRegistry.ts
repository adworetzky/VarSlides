/**
 * useRegistry.ts
 *
 * Custom hook for reading and writing the VarSyncRegistry via Office JS.
 * All component interactions with the registry go through this hook, which
 * keeps the Zustand store and custom XML in sync.
 */

import { useCallback } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { loadRegistry, saveRegistry, cleanOrphanedBindings } from "../lib/registry";
import type { Variable, Binding, VarSyncRegistry } from "../../types";
import { HIGHLIGHT_PALETTE, REGISTRY_VERSION } from "../../types";
import { generateId } from "../lib/linker";

export function useRegistry() {
  const { registry, setRegistry } = useVarSyncStore();

  // ─── Load ──────────────────────────────────────────────────────────────────

  const load = useCallback(async () => {
    const loaded = await loadRegistry();
    setRegistry(loaded);
  }, [setRegistry]);

  // ─── Persist helper ────────────────────────────────────────────────────────

  const persist = useCallback(
    async (next: VarSyncRegistry) => {
      setRegistry(next);
      await saveRegistry(next);
    },
    [setRegistry]
  );

  // ─── Variable CRUD ─────────────────────────────────────────────────────────

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

  // ─── Binding CRUD ──────────────────────────────────────────────────────────

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

  // ─── Cleanup ───────────────────────────────────────────────────────────────

  const runCleanup = useCallback(async () => {
    const clean = cleanOrphanedBindings(registry);
    if (clean.bindings.length !== registry.bindings.length) {
      await persist(clean);
    }
  }, [registry, persist]);

  return {
    registry,
    load,
    addVariable,
    updateVariable,
    deleteVariable,
    addBinding,
    removeBinding,
    updateBinding,
    updateAllBindings,
    runCleanup,
  };
}
