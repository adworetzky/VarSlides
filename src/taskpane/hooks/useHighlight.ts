/**
 * useHighlight.ts
 *
 * Manage highlight mode state — toggling highlights on/off and guarding
 * against highlights persisting to saved files.
 */

import { useCallback, useEffect } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { applyHighlights, stripHighlights } from "../lib/highlighter";
import { saveRegistry } from "../lib/registry";
import type { VarSyncRegistry } from "../../types";

export function useHighlight() {
  const { registry, highlightMode, setHighlightMode, setRegistry } = useVarSyncStore();

  // ─── Toggle ───────────────────────────────────────────────────────────────

  const enableHighlights = useCallback(async () => {
    const updatedRegistry = await applyHighlights(registry);
    setRegistry(updatedRegistry);
    setHighlightMode(true);
  }, [registry, setRegistry, setHighlightMode]);

  const disableHighlights = useCallback(async () => {
    await stripHighlights(registry);
    setHighlightMode(false);
  }, [registry, setHighlightMode]);

  const toggleHighlights = useCallback(async () => {
    if (highlightMode) {
      await disableHighlights();
    } else {
      await enableHighlights();
    }
  }, [highlightMode, enableHighlights, disableHighlights]);

  // ─── Save guard ───────────────────────────────────────────────────────────
  //
  // PowerPoint does not expose a reliable before-save event in Office JS for
  // all platforms. We register a handler where available; on unsupported
  // platforms we warn the user via the UI (see HighlightToggle component).

  useEffect(() => {
    let registered = false;

    const tryRegisterSaveGuard = () => {
      try {
        // DocumentBeforeSave is not universally available; wrap in try/catch
        Office.context.document.addHandlerAsync(
          // @ts-expect-error — DocumentBeforeSave not in all @types/office-js versions
          "documentBeforeSave",
          async (_event: unknown) => {
            if (highlightMode) {
              await stripHighlights(registry);
              setHighlightMode(false);
              // Persist the clean registry so the save writes clean data
              await saveRegistry(registry);
            }
          },
          (result) => {
            registered = result.status === Office.AsyncResultStatus.Succeeded;
          }
        );
      } catch {
        // Save guard not available on this platform; UI warning displayed instead
      }
    };

    tryRegisterSaveGuard();

    return () => {
      if (registered) {
        try {
          Office.context.document.removeHandlerAsync(
            // @ts-expect-error
            "documentBeforeSave",
            {},
            () => {}
          );
        } catch {
          // noop
        }
      }
    };
  }, [highlightMode, registry, setHighlightMode]);

  return {
    highlightMode,
    enableHighlights,
    disableHighlights,
    toggleHighlights,
  };
}
