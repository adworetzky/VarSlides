import type { Binding, Variable, VarSyncRegistry } from "../src/types";
import { REGISTRY_VERSION } from "../src/types";

export function makeVariable(name: string, color?: string): Variable {
  return {
    name,
    value: "$70B",
    color: color ?? "#00C8E8",
    createdAt: new Date().toISOString(),
  };
}

export function makeBinding(overrides: Partial<Binding> = {}): Binding {
  return {
    id: "test-id-1",
    variableName: "AUM",
    shapeId: "shape-1",
    slideIndex: 0,
    paragraphIndex: 0,
    runIndex: 0,
    lastKnownValue: "$70B",
    charOffset: 0,
    originalFontColor: null,
    originalHighlightColor: null,
    ...overrides,
  };
}

export function makeRegistry(
  variables: Variable[] = [],
  bindings: Binding[] = []
): VarSyncRegistry {
  return { version: REGISTRY_VERSION, variables, bindings };
}
