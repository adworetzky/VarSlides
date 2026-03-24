# VarSync

A PowerPoint Office JS task pane add-in that lets you define named variables and link them to text anywhere in a deck — including inline substrings within mixed-content text boxes. Update a variable value once, sync everywhere.

## What it does

- **Variable registry** — create named key/value pairs (`AUM = $70B`) stored in the file's custom XML, so they travel with the `.pptx`
- **Whole-shape linking** — bind a shape's entire text content to a variable
- **Inline linking** — bind a substring inside a mixed text box to a variable (e.g. link just `$70B` inside `"The firm manages $70B across strategies"`)
- **Sync** — one click updates every linked instance across all slides; broken or drifted links are surfaced for manual review
- **Highlight mode** — color-coded overlays show every linked span; stripped automatically before save
- **Broken link recovery** — Re-learn lets you point a broken binding at the correct text without deleting it

## Tech stack

| Layer | Choice |
|---|---|
| Platform | Office JS (PowerPoint) |
| UI | React 18 + TypeScript |
| Build | Vite 5 + `@vitejs/plugin-basic-ssl` (HTTPS required for sideloading) |
| State | Zustand |
| Styling | Tailwind CSS (no Fluent UI) |
| Storage | PowerPoint custom XML parts |
| Tests | Jest + `office-addin-mock` |

## Getting started

```bash
npm install
npm run dev        # starts HTTPS dev server on https://localhost:3000
```

Sideload in PowerPoint: **Insert → Add-ins → Upload My Add-in → select `manifest.xml`**

```bash
npm test           # run unit tests (no PowerPoint host required)
npm run build      # production build to dist/
```

## Project structure

```
varsync/
├── manifest.xml                        # add-in manifest for sideloading
├── NOTES.md                            # Script Lab findings, API behavior notes
├── src/
│   ├── types/index.ts                  # all shared TypeScript interfaces
│   └── taskpane/
│       ├── index.html + index.tsx      # entry point (Office.onReady → React)
│       ├── App.tsx                     # 4-tab shell + Sync All header
│       ├── store/useVarSyncStore.ts    # Zustand store
│       ├── components/
│       │   ├── VariablesPanel.tsx      # variable CRUD + per-variable sync
│       │   ├── LinkPanel.tsx           # shape/selection linking UI
│       │   ├── BindingsPanel.tsx       # binding list + unlink
│       │   ├── HighlightToggle.tsx     # highlight mode + color legend
│       │   ├── SyncSummaryPanel.tsx    # post-sync result display
│       │   └── BrokenLinkRow.tsx       # broken binding with Re-learn / Unlink
│       ├── hooks/
│       │   ├── useRegistry.ts          # CRUD bridge: Zustand ↔ custom XML
│       │   ├── useSelection.ts         # tracks active PowerPoint selection
│       │   └── useHighlight.ts         # highlight toggle + save guard
│       └── lib/                        # pure logic, fully unit-testable
│           ├── registry.ts             # XML serialize/deserialize, cleanOrphanedBindings
│           ├── linker.ts               # run splitting, binding creation
│           ├── syncer.ts               # Clean/Recoverable/Broken classification + sync
│           └── highlighter.ts          # color assignment, apply/strip highlights
└── tests/
    ├── fixtures.ts                     # shared makeVariable / makeBinding / makeRegistry
    ├── registry.test.ts
    ├── linker.test.ts
    ├── syncer.test.ts
    └── highlighter.test.ts
```

## Binding states

When syncing, each binding is classified:

| State | Meaning | Action |
|---|---|---|
| **Clean** | Run at `runIndex` still contains `lastKnownValue` | Update in place |
| **Recoverable** | Run was remerged but value found in paragraph text | Auto-repair: re-split, update |
| **Broken** | `lastKnownValue` not found anywhere in paragraph | Surface to user; Re-learn or Unlink |

## Highlight mode

- Each variable gets a color from an 8-color palette; colors cycle if more than 8 variables exist
- Highlights are applied to run font `highlightColor` in Office JS
- **Highlights must never persist to saved files.** A `documentBeforeSave` handler strips them automatically where the API is available (Windows desktop). On Mac/Web, the UI shows a persistent warning while highlights are active.
- Original `fontColor` and `highlightColor` are stored in each binding record for clean restoration

## Custom XML schema

All data stored as a single JSON blob under namespace `http://varsync/registry/v1`:

```typescript
interface VarSyncRegistry {
  version: string;
  variables: Variable[];   // { name, value, color, createdAt }
  bindings: Binding[];     // { id, variableName, shapeId, slideIndex,
                           //   paragraphIndex, runIndex, lastKnownValue,
                           //   charOffset, originalFontColor, originalHighlightColor }
}
```

## Phase 3 (inline linking) — pre-requisite

Run the Script Lab validation in `NOTES.md` on both Windows and Mac before deploying inline substring linking. The `getSelectedTextRange()` and `paragraph.textRange.runs` APIs have known behavioral differences between platforms.

## Known limitations (v1)

- Charts, images, SmartArt, and tables cannot be linked — an error is shown if attempted
- Auto-sync on value edit is a v2 feature
- Undo/redo beyond native PowerPoint undo is out of scope
