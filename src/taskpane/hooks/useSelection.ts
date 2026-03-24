/**
 * useSelection.ts
 *
 * Watch the active PowerPoint selection and expose a SelectionContext.
 *
 * NOTE: PowerPoint Office JS selection events fire on shape selection changes.
 * Text-level selection within a shape is queryable via the active selection
 * API but does NOT fire a selection-changed event — the hook polls on demand
 * and relies on the user clicking "Link Selection" to capture the text selection.
 *
 * Platform notes (see NOTES.md):
 *   - Windows desktop: PowerPoint.TextRange selection works reliably.
 *   - Mac / Web: getSelectedDataAsync may be needed as a fallback.
 */

import { useCallback, useEffect } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import type { SelectionContext } from "../../types";

export function useSelection() {
  const { selectionContext, setSelectionContext } = useVarSyncStore();

  // ─── Read current selection ──────────────────────────────────────────────

  const refreshSelection = useCallback(async () => {
    try {
      await PowerPoint.run(async (context) => {
        // Attempt to get the selected shapes
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (!selectedShapes.items || selectedShapes.items.length === 0) {
          setSelectionContext(null);
          return;
        }

        const shape = selectedShapes.items[0];
        shape.load("id, name");

        // Determine which slide this shape is on
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        await context.sync(); // second sync for shape properties

        // Find slide index by searching shapes
        let slideIndex = 0;
        for (let i = 0; i < slides.items.length; i++) {
          const slideShapes = slides.items[i].shapes;
          slideShapes.load("items");
          await context.sync();
          const found = slideShapes.items.find((s) => s.id === shape.id);
          if (found) {
            slideIndex = i;
            break;
          }
        }

        // Check for text selection
        let hasTextSelection = false;
        try {
          const selectedRange = context.presentation.getSelectedTextRange();
          selectedRange.load("text");
          await context.sync();
          hasTextSelection =
            typeof selectedRange.text === "string" && selectedRange.text.length > 0;
        } catch {
          // No text selection active
          hasTextSelection = false;
        }

        const ctx: SelectionContext = {
          shapeId: shape.id,
          shapeName: shape.name,
          slideIndex,
          hasTextSelection,
        };
        setSelectionContext(ctx);
      });
    } catch {
      setSelectionContext(null);
    }
  }, [setSelectionContext]);

  // ─── Subscribe to selection change events ────────────────────────────────

  useEffect(() => {
    let handler: Office.EventHandlerResult | null = null;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        void refreshSelection();
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          handler = result.value as Office.EventHandlerResult;
        }
      }
    );

    // Initial read
    void refreshSelection();

    return () => {
      if (handler) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          {},
          () => {}
        );
      }
    };
  }, [refreshSelection]);

  return { selectionContext, refreshSelection };
}
