import { generateId, isLinkableShapeType } from "../src/taskpane/lib/linker";

describe("generateId", () => {
  it("generates a non-empty string", () => {
    const id = generateId();
    expect(typeof id).toBe("string");
    expect(id.length).toBeGreaterThan(0);
  });

  it("generates unique IDs on successive calls", () => {
    const ids = new Set(Array.from({ length: 100 }, () => generateId()));
    expect(ids.size).toBe(100);
  });

  it("generates IDs in UUID-like format", () => {
    const id = generateId();
    expect(id).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i
    );
  });
});

describe("isLinkableShapeType", () => {
  it("returns true for a plain text box shape", () => {
    expect(isLinkableShapeType("TextBox")).toBe(true);
  });

  it("returns true for AutoShape", () => {
    expect(isLinkableShapeType("AutoShape")).toBe(true);
  });

  it("returns false for Chart", () => {
    expect(isLinkableShapeType("Chart")).toBe(false);
  });

  it("returns false for Picture", () => {
    expect(isLinkableShapeType("Picture")).toBe(false);
  });

  it("returns false for SmartArt", () => {
    expect(isLinkableShapeType("SmartArt")).toBe(false);
  });

  it("returns false for Table", () => {
    expect(isLinkableShapeType("Table")).toBe(false);
  });

  it("returns false for Media", () => {
    expect(isLinkableShapeType("Media")).toBe(false);
  });

  it("returns false for 3DModel", () => {
    expect(isLinkableShapeType("3DModel")).toBe(false);
  });

  it("returns true for unknown/custom shape types (allow by default)", () => {
    expect(isLinkableShapeType("CustomShape")).toBe(true);
    expect(isLinkableShapeType("PlaceHolder")).toBe(true);
  });
});
