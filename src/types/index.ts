// ─────────────────────────────────────────────────────────────────────────────
// VarSync — shared TypeScript interfaces
// ─────────────────────────────────────────────────────────────────────────────

export interface Variable {
  name: string; // unique key, user-defined
  value: string; // current value
  color: string; // hex color for highlight mode
  createdAt: string; // ISO timestamp
}

export interface Binding {
  id: string; // UUID
  variableName: string; // key into variable registry
  shapeId: string; // PowerPoint shape ID
  shapeName?: string; // human-readable shape name (stored at link time for display)
  slideIndex: number; // 0-based slide index
  paragraphIndex: number; // index within shape's text body
  runIndex: number; // index of the specific run after splitting
  lastKnownValue: string; // value at time of linking, used for recovery
  charOffset: number; // character position within paragraph at time of linking
  originalFontColor: string | null; // stored at link time for highlight restore
  originalHighlightColor: string | null;
}

export interface BrokenBinding extends Binding {
  slideNumber: number; // 1-based for display
  shapeName: string;
}

export interface VarSyncRegistry {
  version: string;
  variables: Variable[];
  bindings: Binding[];
}

export interface SelectionContext {
  shapeId: string;
  shapeName: string;
  slideIndex: number;
  hasTextSelection: boolean;
}

export interface SyncSummary {
  clean: number;
  recovered: number;
  broken: BrokenBinding[];
}

export type BindingState = "clean" | "recoverable" | "broken";

export interface ClassifiedBinding {
  binding: Binding;
  state: BindingState;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

export const HIGHLIGHT_PALETTE = [
  "#00C8E8", // cyan
  "#E8A020", // amber
  "#A8FF00", // lime
  "#FF3B3B", // red
  "#A78BFA", // purple
  "#FB923C", // orange
  "#F472B6", // pink
  "#34D399", // green
] as const;

export const REGISTRY_NAMESPACE = "http://varsync/registry/v1";
export const REGISTRY_VERSION = "1.0.0";
