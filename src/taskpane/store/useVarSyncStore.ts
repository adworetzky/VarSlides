import { create } from "zustand";
import type {
  VarSyncRegistry,
  SelectionContext,
  SyncSummary,
} from "../../types";
import { REGISTRY_VERSION } from "../../types";

interface VarSyncStore {
  registry: VarSyncRegistry;
  highlightMode: boolean;
  selectionContext: SelectionContext | null;
  syncSummary: SyncSummary | null;
  // actions
  setRegistry: (r: VarSyncRegistry) => void;
  setHighlightMode: (on: boolean) => void;
  setSelectionContext: (ctx: SelectionContext | null) => void;
  setSyncSummary: (s: SyncSummary | null) => void;
}

const emptyRegistry = (): VarSyncRegistry => ({
  version: REGISTRY_VERSION,
  variables: [],
  bindings: [],
});

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
