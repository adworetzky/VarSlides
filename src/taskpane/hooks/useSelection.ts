// NOTE: DocumentSelectionChanged fires on shape selection changes but NOT on
// text selection changes within a shape. Text selection is captured on-demand
// when the user clicks "Link Selection". See NOTES.md for platform details.

import { useCallback, useEffect } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import type { SelectionContext } from "../../types";

export function useSelection() {
  const { selectionContext, setSelectionContext } = useVarSyncStore();

  const refreshSelection = useCallback(async () => {
    try {
      await PowerPoint.run(async (context) => {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (!selectedShapes.items || selectedShapes.items.length === 0) {
          setSelectionContext(null);
          return;
        }

        const shape = selectedShapes.items[0];
        shape.load("id, name");

        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync(); // resolves shape properties + slides in one round-trip

        // Determine which slide this shape is on by scanning all slides' shapes.
        // Batch-load all slides' shape collections before syncing.
        for (const slide of slides.items) {
          slide.shapes.load("items");
        }
        await context.sync();

        let slideIndex = 0;
        for (let i = 0; i < slides.items.length; i++) {
          if (slides.items[i].shapes.items.find((s) => s.id === shape.id)) {
            slideIndex = i;
            break;
          }
        }

        let hasTextSelection = false;
        try {
          const selectedRange = context.presentation.getSelectedTextRange();
          selectedRange.load("text");
          await context.sync();
          hasTextSelection =
            typeof selectedRange.text === "string" && selectedRange.text.length > 0;
        } catch {
          hasTextSelection = false;
        }

        setSelectionContext({
          shapeId: shape.id,
          shapeName: shape.name,
          slideIndex,
          hasTextSelection,
        });
      });
    } catch {
      setSelectionContext(null);
    }
  }, [setSelectionContext]);

  useEffect(() => {
    let handler: Office.EventHandlerResult | null = null;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => { void refreshSelection(); },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          handler = result.value as Office.EventHandlerResult;
        }
      }
    );

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
