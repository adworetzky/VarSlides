import { matchSnippet } from "../src/taskpane/lib/finder";

describe("matchSnippet", () => {
  it("returns the full text when match fits within context window", () => {
    const result = matchSnippet("Hello world", "world", 6);
    expect(result).toEqual({ before: "Hello ", match: "world", after: "" });
  });

  it("truncates long before text with ellipsis", () => {
    const longText = "abcdefghijklmnopqrstuvwxyz$70B more text";
    const result = matchSnippet(longText, "$70B", 26, 10);
    expect(result.before).toMatch(/^…/);
    expect(result.match).toBe("$70B");
  });

  it("truncates long after text with ellipsis", () => {
    const longText = "$70B abcdefghijklmnopqrstuvwxyz";
    const result = matchSnippet(longText, "$70B", 0, 10);
    expect(result.after).toMatch(/…$/);
    expect(result.match).toBe("$70B");
  });

  it("no ellipsis when text fits exactly in context", () => {
    const result = matchSnippet("abc$70Bxyz", "$70B", 3, 3);
    expect(result).toEqual({ before: "abc", match: "$70B", after: "xyz" });
  });

  it("handles match at position 0", () => {
    const result = matchSnippet("$70B manages assets", "$70B", 0);
    expect(result.before).toBe("");
    expect(result.match).toBe("$70B");
    expect(result.after).toBe(" manages assets");
  });

  it("handles match at end of string", () => {
    const result = matchSnippet("firm manages $70B", "$70B", 13);
    expect(result.before).toBe("firm manages ");
    expect(result.match).toBe("$70B");
    expect(result.after).toBe("");
  });

  it("handles single-character match", () => {
    const result = matchSnippet("abc", "b", 1, 1);
    expect(result).toEqual({ before: "a", match: "b", after: "c" });
  });

  it("handles match spanning the entire string", () => {
    const result = matchSnippet("$70B", "$70B", 0);
    expect(result).toEqual({ before: "", match: "$70B", after: "" });
  });
});
