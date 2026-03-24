# VarSync — API Behavior Notes & Script Lab Findings

## Purpose

This file documents Office JS API behavior findings from Script Lab validation sessions,
platform differences, and any constraints discovered during development. **Phase 3
(inline/substring linking) must not be deployed without completing the Script Lab
validation described below.**

---

## Phase 3 Pre-Requisite: Script Lab Validation

### What to validate

Before implementing run-level text selection in `linker.ts`, confirm the following in
Script Lab on **both Windows desktop and Mac**:

1. **`context.presentation.getSelectedTextRange()` from task pane context**
   - While a user has text selected inside a shape (in text editing mode), can the
     task pane JavaScript read the selected range?
   - Specifically: does `.text` return the selected substring, or does it always return
     the full shape text?

2. **`paragraph.textRange.runs` iteration**
   - Can individual runs be loaded and their `.text` property read without error?
   - Does PowerPoint preserve run boundaries after a text rewrite, or does it remerge
     adjacent same-formatted runs?

3. **`paragraph.textRange.getSubstring(start, length)`**
   - Is `getSubstring` available in the PowerPoint JS API?
   - Does it behave as expected for character-level targeting?

4. **Run count after `textRange.text = newText` assignment**
   - After assigning a new text string to a paragraph's `textRange.text`, how many
     runs does PowerPoint create?
   - Does it flatten to one run, or preserve the structure?

### Script Lab snippets

```javascript
// Snippet 1: Read selected text from task pane
await PowerPoint.run(async (context) => {
  const range = context.presentation.getSelectedTextRange();
  range.load("text");
  await context.sync();
  console.log("Selected text:", range.text);
});
```

```javascript
// Snippet 2: Enumerate runs in a paragraph
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const shape = slides.items[0].shapes.items[0];
  const para = shape.textFrame.textRange.paragraphs.items[0];
  para.textRange.runs.load("items");
  await context.sync();
  para.textRange.runs.items.forEach((run, i) => {
    run.textRange.load("text");
  });
  await context.sync();
  para.textRange.runs.items.forEach((run, i) => {
    console.log(`Run ${i}:`, run.textRange.text);
  });
});
```

```javascript
// Snippet 3: getSubstring
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const shape = slides.items[0].shapes.items[0];
  const range = shape.textFrame.textRange;
  range.load("text");
  await context.sync();
  const sub = range.getSubstring(0, 5);
  sub.load("text");
  await context.sync();
  console.log("Substring:", sub.text);
});
```

---

## Findings Log

> Update this section as testing is completed. Prefix each entry with the date and platform.

### [DATE] Windows Desktop — Office Version X.X

- [ ] `getSelectedTextRange()` reads selection from task pane: **TBD**
- [ ] `paragraph.textRange.runs` iteration works: **TBD**
- [ ] `getSubstring` available and correct: **TBD**
- [ ] Run count after text rewrite: **TBD**

### [DATE] Mac — Office Version X.X

- [ ] `getSelectedTextRange()` reads selection from task pane: **TBD**
- [ ] `paragraph.textRange.runs` iteration works: **TBD**
- [ ] `getSubstring` available and correct: **TBD**
- [ ] Run count after text rewrite: **TBD**

---

## Known API Gaps and Workarounds

### `DocumentBeforeSave` event

- Available: Windows Desktop (limited support), **not available** on Mac or Web.
- Workaround: Display a prominent warning in the `HighlightToggle` component when
  highlights are active and the user has not yet disabled them. The warning persists
  until highlights are toggled off.

### Selection events for text ranges

- `Office.EventType.DocumentSelectionChanged` fires on shape selection changes but
  does NOT fire when the user changes text selection within an already-selected shape.
- Workaround: Text selection is captured on-demand when the user clicks "Link Selection".
  `refreshSelection()` is called at that point to read the current state.

### Run remerging

- PowerPoint may remerge adjacent runs with identical formatting when the user edits
  nearby text or when the document is saved and reopened.
- Handled by the "Recoverable" state in `syncer.ts` — `lastKnownValue` string search
  falls back to character offset when the run index is stale.

### `getSelectedShapes()` API

- Available in PowerPoint JS API 1.5+.
- Falls back gracefully to `null` selection context if the API is not available.

---

## Build Notes

- HTTPS dev server is required for Office JS sideloading (self-signed cert via
  `@vitejs/plugin-basic-ssl` is sufficient for local development).
- Custom XML parts are scoped to the presentation file. Opening the file on a
  different machine preserves all variables and bindings.
- The registry namespace `http://varsync/registry/v1` must not be changed between
  versions without a migration path.
