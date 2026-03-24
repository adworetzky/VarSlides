import {
  classifyBinding,
  findLastKnownValueOffset,
  reLearnBinding,
} from "../src/taskpane/lib/syncer";
import { makeBinding } from "./fixtures";

describe("classifyBinding", () => {
  it('returns "clean" when run at runIndex contains lastKnownValue', () => {
    const binding = makeBinding({ runIndex: 1, lastKnownValue: "$70B" });
    const runTexts = ["The firm's AUM is ", "$70B", " across all strategies"];
    const result = classifyBinding(binding, runTexts, runTexts.join(""));
    expect(result).toBe("clean");
  });

  it('returns "recoverable" when run is remerged but value found in paragraph', () => {
    const binding = makeBinding({ runIndex: 5, lastKnownValue: "$70B" });
    // runIndex 5 doesn't exist in runTexts, but value is present in paragraph
    const runTexts = ["The firm's AUM is $70B across all strategies"];
    const paraText = "The firm's AUM is $70B across all strategies";
    const result = classifyBinding(binding, runTexts, paraText);
    expect(result).toBe("recoverable");
  });

  it('returns "recoverable" when lastKnownValue found in paragraph regardless of run structure', () => {
    const binding = makeBinding({ runIndex: 0, lastKnownValue: "$70B", charOffset: 0 });
    const runTexts = ["Something else entirely"];
    const paraText = "The firm's AUM is $70B across strategies";
    const result = classifyBinding(binding, runTexts, paraText);
    expect(result).toBe("recoverable");
  });

  it('returns "broken" when lastKnownValue not found in paragraph', () => {
    const binding = makeBinding({ lastKnownValue: "$70B" });
    const runTexts = ["Updated text that no longer has it"];
    const paraText = "Updated text that no longer has it";
    const result = classifyBinding(binding, runTexts, paraText);
    expect(result).toBe("broken");
  });

  it('returns "broken" when runTexts is empty and paragraph does not contain lastKnownValue', () => {
    const binding = makeBinding({ lastKnownValue: "$100M" });
    const result = classifyBinding(binding, [], "The firm's AUM is $70B");
    expect(result).toBe("broken");
  });

  it('returns "clean" for exact run match even with multiple candidates', () => {
    const binding = makeBinding({ runIndex: 2, lastKnownValue: "$70B" });
    const runTexts = ["Prefix ", "Other ", "$70B", " suffix"];
    const result = classifyBinding(binding, runTexts, runTexts.join(""));
    expect(result).toBe("clean");
  });
});

describe("findLastKnownValueOffset", () => {
  it("returns correct offset when value is present", () => {
    const para = "The firm's AUM is $70B across strategies";
    expect(findLastKnownValueOffset("$70B", para)).toBe(18);
  });

  it("returns -1 when value is not present", () => {
    const para = "No variable here";
    expect(findLastKnownValueOffset("$70B", para)).toBe(-1);
  });

  it("returns index of first occurrence", () => {
    const para = "$70B and then $70B again";
    expect(findLastKnownValueOffset("$70B", para)).toBe(0);
  });

  it("handles empty lastKnownValue", () => {
    // empty string is always found at index 0
    expect(findLastKnownValueOffset("", "any text")).toBe(0);
  });
});

describe("reLearnBinding", () => {
  it("updates runIndex, charOffset, and lastKnownValue", () => {
    const original = makeBinding({
      runIndex: 0,
      charOffset: 0,
      lastKnownValue: "$70B",
    });
    const updated = reLearnBinding(original, 3, 22, "$85B");
    expect(updated.runIndex).toBe(3);
    expect(updated.charOffset).toBe(22);
    expect(updated.lastKnownValue).toBe("$85B");
  });

  it("preserves all other binding fields", () => {
    const original = makeBinding();
    const updated = reLearnBinding(original, 1, 5, "new");
    expect(updated.id).toBe(original.id);
    expect(updated.variableName).toBe(original.variableName);
    expect(updated.shapeId).toBe(original.shapeId);
    expect(updated.slideIndex).toBe(original.slideIndex);
  });

  it("does not mutate the original binding", () => {
    const original = makeBinding({ runIndex: 0 });
    reLearnBinding(original, 5, 10, "new");
    expect(original.runIndex).toBe(0);
  });
});
