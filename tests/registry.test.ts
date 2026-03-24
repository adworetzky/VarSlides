import { cleanOrphanedBindings } from "../src/taskpane/lib/registry";
import { makeVariable, makeBinding, makeRegistry } from "./fixtures";

describe("cleanOrphanedBindings", () => {
  it("keeps bindings whose variable still exists", () => {
    const registry = makeRegistry(
      [makeVariable("AUM")],
      [makeBinding({ variableName: "AUM" })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(1);
  });

  it("removes bindings whose variable has been deleted", () => {
    const registry = makeRegistry(
      [], // no variables
      [makeBinding({ variableName: "AUM" })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(0);
  });

  it("removes bindings with missing id", () => {
    const registry = makeRegistry(
      [makeVariable("AUM")],
      [makeBinding({ id: "" })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(0);
  });

  it("removes bindings with missing shapeId", () => {
    const registry = makeRegistry(
      [makeVariable("AUM")],
      [makeBinding({ shapeId: "" })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(0);
  });

  it("removes bindings with negative slideIndex", () => {
    const registry = makeRegistry(
      [makeVariable("AUM")],
      [makeBinding({ slideIndex: -1 })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(0);
  });

  it("removes bindings with negative runIndex", () => {
    const registry = makeRegistry(
      [makeVariable("AUM")],
      [makeBinding({ runIndex: -1 })]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(0);
  });

  it("preserves multiple valid bindings across variables", () => {
    const registry = makeRegistry(
      [makeVariable("AUM"), makeVariable("YEAR")],
      [
        makeBinding({ id: "b1", variableName: "AUM" }),
        makeBinding({ id: "b2", variableName: "YEAR" }),
        makeBinding({ id: "b3", variableName: "GONE" }), // orphan
      ]
    );
    const clean = cleanOrphanedBindings(registry);
    expect(clean.bindings).toHaveLength(2);
    expect(clean.bindings.map((b) => b.id)).toEqual(["b1", "b2"]);
  });

  it("does not mutate the original registry", () => {
    const bindings = [makeBinding({ variableName: "DELETED" })];
    const registry = makeRegistry([], bindings);
    cleanOrphanedBindings(registry);
    expect(registry.bindings).toHaveLength(1); // original unchanged
  });
});
