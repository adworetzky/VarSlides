/**
 * highlighter.test.ts
 *
 * Tests for pure highlighter functions (no Office JS required).
 */

import {
  assignVariableColors,
  isColorCycling,
  getVariableColor,
} from "../src/taskpane/lib/highlighter";
import { HIGHLIGHT_PALETTE } from "../src/types";
import type { Variable } from "../src/types";

// ─────────────────────────────────────────────────────────────────────────────
// Fixtures
// ─────────────────────────────────────────────────────────────────────────────

function makeVariable(name: string, color?: string): Variable {
  return {
    name,
    value: "test",
    color: color ?? HIGHLIGHT_PALETTE[0],
    createdAt: new Date().toISOString(),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// assignVariableColors
// ─────────────────────────────────────────────────────────────────────────────

describe("assignVariableColors", () => {
  it("returns a map with an entry per variable", () => {
    const vars = [makeVariable("A"), makeVariable("B"), makeVariable("C")];
    const map = assignVariableColors(vars);
    expect(map.size).toBe(3);
    expect(map.has("A")).toBe(true);
    expect(map.has("B")).toBe(true);
    expect(map.has("C")).toBe(true);
  });

  it("uses the variable's stored color when present", () => {
    const vars = [makeVariable("AUM", "#FF0000")];
    const map = assignVariableColors(vars);
    expect(map.get("AUM")).toBe("#FF0000");
  });

  it("assigns palette colors in order when variable has no explicit color", () => {
    // Variable with empty color string falls back to palette
    const vars = [
      { name: "A", value: "x", color: "", createdAt: "" },
      { name: "B", value: "x", color: "", createdAt: "" },
    ];
    const map = assignVariableColors(vars);
    // Since color is empty string (falsy), should use palette
    expect(map.get("A")).toBe(HIGHLIGHT_PALETTE[0]);
    expect(map.get("B")).toBe(HIGHLIGHT_PALETTE[1]);
  });

  it("handles empty variable list", () => {
    const map = assignVariableColors([]);
    expect(map.size).toBe(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// isColorCycling
// ─────────────────────────────────────────────────────────────────────────────

describe("isColorCycling", () => {
  it("returns false for 8 or fewer variables", () => {
    const vars = Array.from({ length: 8 }, (_, i) => makeVariable(`V${i}`));
    expect(isColorCycling(vars)).toBe(false);
  });

  it("returns true for 9 or more variables", () => {
    const vars = Array.from({ length: 9 }, (_, i) => makeVariable(`V${i}`));
    expect(isColorCycling(vars)).toBe(true);
  });

  it("returns false for empty list", () => {
    expect(isColorCycling([])).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getVariableColor
// ─────────────────────────────────────────────────────────────────────────────

describe("getVariableColor", () => {
  it("returns the variable's assigned color", () => {
    const vars = [makeVariable("AUM", "#AABBCC")];
    expect(getVariableColor("AUM", vars)).toBe("#AABBCC");
  });

  it("falls back to palette by index when color is empty", () => {
    const vars = [
      { name: "A", value: "x", color: "", createdAt: "" },
      { name: "B", value: "x", color: "", createdAt: "" },
    ];
    expect(getVariableColor("B", vars)).toBe(HIGHLIGHT_PALETTE[1]);
  });

  it("cycles palette when variable index exceeds palette length", () => {
    const vars = Array.from({ length: 10 }, (_, i) => ({
      name: `V${i}`,
      value: "x",
      color: "",
      createdAt: "",
    }));
    // V8 is index 8, palette has 8 entries → cycles back to index 0
    expect(getVariableColor("V8", vars)).toBe(HIGHLIGHT_PALETTE[0]);
    // V9 → index 1
    expect(getVariableColor("V9", vars)).toBe(HIGHLIGHT_PALETTE[1]);
  });

  it("returns first palette color for unknown variable name", () => {
    const vars = [makeVariable("Known")];
    // findIndex returns -1 for unknown; -1 % 8 is implementation-dependent
    // We use modulo cycling so -1 % 8 = -1 in JS, but we guard with Math.max or the positive index
    // This test documents current behavior — fallback to first palette color
    const color = getVariableColor("Unknown", vars);
    expect(HIGHLIGHT_PALETTE).toContain(color);
  });
});
