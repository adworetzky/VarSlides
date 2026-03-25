# CLAUDE.md — VarSync development guide

## Commands

```bash
npm run dev        # HTTPS dev server (required for Office JS sideloading)
npm test           # Jest unit tests — run after every change
npm run build      # production build
```

Tests must pass before committing. All logic in `src/taskpane/lib/` is pure and runs without a PowerPoint host.

## Architecture

### Data flow

```
PowerPoint host
      ↕ Office JS
  lib/ (registry.ts, linker.ts, syncer.ts, highlighter.ts, finder.ts, navigator.ts)
      ↕ async functions
  hooks/ (useRegistry, useSelection, useHighlight)
      ↕ Zustand
  components/ → reads from store, calls hooks
```

**Registry writes always go: Zustand first → custom XML second.** The `persist()` helper in `useRegistry` does an optimistic store update and rolls back if the XML write fails. Never write directly to custom XML from a component.

### Store shape

```typescript
{ registry, highlightMode, selectionContext, syncSummary }
```

`registry` is the single source of truth for variables and bindings. Everything else is derived or ephemeral UI state.

### Custom XML

Namespace: `http://varsync/registry/v1`. A single XML part holds the full JSON registry. `saveRegistry` deletes all existing parts under the namespace then adds a new one — in-place update is not possible with the Office JS custom XML API.

## Key constraints

### Never use Fluent UI

UI must be built with Tailwind CSS utility classes only. Task pane is 320px wide — design compact and data-dense.

### Highlights must not persist to saved files

`useHighlight` registers a `documentBeforeSave` handler (where available) that strips highlights before the save completes. This handler is registered **once on mount** and reads live store state via `useVarSyncStore.getState()` at event time — do not add `registry` or `highlightMode` to the effect dependency array, as that would re-register on every state change.

On Mac/Web where `documentBeforeSave` is unavailable, `HighlightToggle` shows a persistent warning. Do not remove this warning.

### Phase 3 inline linking — Script Lab first

Do not implement or modify run-splitting logic in `linker.ts` without first completing the Script Lab validation documented in `NOTES.md`. The `getSelectedTextRange()` and `paragraph.textRange.runs` APIs behave differently on Windows vs Mac. Findings must be recorded in `NOTES.md` before proceeding.

## Office JS patterns

### Batching context.sync()

Load all properties before syncing — never `await context.sync()` inside a loop:

```typescript
// ✓ correct — one round-trip
for (const run of runs.items) {
  run.textRange.load("text");
}
await context.sync();
for (const run of runs.items) {
  process(run.textRange.text);
}

// ✗ wrong — N round-trips
for (const run of runs.items) {
  run.textRange.load("text");
  await context.sync();           // <-- N+1
  process(run.textRange.text);
}
```

### Slide shape lookup

When searching for a shape across all slides, batch-load all slide shape collections before syncing:

```typescript
for (const slide of slides.items) {
  slide.shapes.load("items");
}
await context.sync();
// now iterate synchronously
```

### emptyRegistry()

The canonical empty registry is exported from `src/taskpane/lib/registry.ts`. Do not define it elsewhere.

## Testing

### What is unit-testable without Office JS

Everything in `src/taskpane/lib/`:
- `registry.ts` — `cleanOrphanedBindings`, `emptyRegistry`, XML helpers
- `linker.ts` — `generateId`, `isLinkableShapeType`
- `syncer.ts` — `classifyBinding`, `findLastKnownValueOffset`, `reLearnBinding`
- `highlighter.ts` — `assignVariableColors`, `isColorCycling`, `getVariableColor`
- `finder.ts` — `matchSnippet`

Office JS calls (`PowerPoint.run`, `Office.context`) live only in the async functions that take a `context` parameter. Keep this separation.

### Shared fixtures

Use `tests/fixtures.ts` for `makeVariable`, `makeBinding`, and `makeRegistry`. Do not add local copies to individual test files.

### Office global stub

`tests/setupOffice.ts` provides minimal `Office` and `PowerPoint` globals. Extend it if new Office APIs need to be called from lib functions that gain unit test coverage.

## Binding classification logic

`classifyBinding` in `syncer.ts`:

1. If `runTexts[binding.runIndex] === lastKnownValue` → **clean** (run is intact)
2. Else if `paragraphText.includes(lastKnownValue)` → **recoverable** (run remerged, value still present)
3. Else → **broken** (value edited or deleted on slide)

Recoverable bindings are auto-repaired during sync via `findLastKnownValueOffset` + `getSubstring`. Broken bindings are never silently overwritten — they surface in `SyncSummaryPanel` for user action.

## UX conventions

### Terminology
The app uses **"link"** (verb) and **"binding"** (noun) consistently across the UI. Do not mix in "connect", "attach", or "reference".

### Loading states
Use `"…"` (ellipsis character, not three dots) for inline loading indicators. For full-button loading states, use descriptive text like `"Syncing…"`, `"Linking…"`, `"Scanning…"`.

### Error messages
All error messages must include a dismiss button (`×`). Use red-950/40 background, red-800/50 border, red-400 text.

### Health badges
Per-variable health is derived from `registry.bindings` — do not store it in the Zustand store. Four states, checked in priority order:
- **"✓ synced"** (green) — all bindings have `lastKnownValue === variable.value`
- **"⟳ stale"** (sky) — all bindings have a non-empty `lastKnownValue`, but at least one differs from the current `variable.value` (changed since last sync)
- **"! broken"** (amber) — at least one binding has empty `lastKnownValue` (structurally broken)
- **"— no links"** (gray) — variable has no bindings yet

`Sync All` button in the header glows and shows a `●` dot whenever any binding is stale or broken.

### SyncSummaryPanel auto-dismiss
The panel auto-dismisses after 3 s when there are no broken bindings. When broken bindings exist, it stays until dismissed.

### shapeName on Binding
`Binding.shapeName` is optional (`string | undefined`). It is stored at link time for display purposes only — it is **not** used for any lookup or sync logic. Always prefer `shapeName` over raw `shapeId` in the UI, falling back to `shapeId.slice(0, 12)` for old bindings without it.

### BrokenBinding.currentText
`BrokenBinding.currentText?: string` is the paragraph text captured from the slide at sync time, giving users a "Expected X, Found Y" diff view without a second Office JS round-trip. It is only populated by `syncVariable()` for bindings classified as broken at runtime — it is **not** present on structurally broken bindings surfaced from the store (those have an empty `lastKnownValue` and no text to diff against).

### Inline occurrence list in VariablesPanel
Each variable row has a `N links ▾` toggle that expands a compact list of its current bindings (slide number, shape name, ✓/⟳ per-binding health). Clicking a slide number calls `navigateToBinding()`. This list reads directly from `registry.bindings` — it is derived state, not fetched from Office JS. The "Find" button still opens `FindLinkPanel` for searching new unlinked occurrences; the two are complementary.

## Color palette

8 fixed colors in `HIGHLIGHT_PALETTE` (`src/types/index.ts`). Variable colors are assigned by position index and cycle if more than 8 variables exist. The `HighlightToggle` legend shows a cycling warning when active.

## Things that will break

- **Removing `REGISTRY_NAMESPACE`** — all existing custom XML parts in user files will become unreadable. Add a migration path if the namespace must change.
- **Changing `Binding` fields** — `lastKnownValue` and `charOffset` are load-bearing for recovery. Renaming them requires migrating stored JSON. `BrokenBinding.currentText` is ephemeral (not persisted) — it is safe to rename or remove.
- **Adding `registry` or `highlightMode` to the `useHighlight` save-guard `useEffect` deps** — this re-registers the handler on every state change, which is expensive and can cause double-strip on save.
- **Calling `saveRegistry` directly from a component** — bypasses Zustand, creating a store/XML split-brain. Always go through `useRegistry`.
- **Storing health state in Zustand** — health is cheap to derive from bindings on every render. A cached store field would go stale and require manual invalidation.
