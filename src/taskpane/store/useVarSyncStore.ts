import { create } from "zustand";
import type { VarSyncRegistry, SelectionContext, SyncSummary } from "../../types";
import { emptyRegistry } from "../lib/registry";

interface VarSyncStore {
  registry: VarSyncRegistry;
  highlightMode: boolean;
  selectionContext: SelectionContext | null;
  syncSummary: SyncSummary | null;
  setRegistry: (r: VarSyncRegistry) => void;
  setHighlightMode: (on: boolean) => void;
  setSelectionContext: (ctx: SelectionContext | null) => void;
  setSyncSummary: (s: SyncSummary | null) => void;
}

export const useVarSyncStore = create<VarSyncStore>((set) => ({
  registry: emptyRegistry(),
  highlightMode: false,
  selectionContext: null,
  syncSummary: null,

  setRegistry: (r) => set({ registry: r }),
  setHighlightMode: (on) => set({ highlightMode: on }),
  setSelectionContext: (ctx) => set({ selectionContext: ctx }),
  setSyncSummary: (s) => set({ syncSummary: s }),
}));
