import { useCallback, useEffect } from "react";
import { useVarSyncStore } from "../store/useVarSyncStore";
import { applyHighlights, stripHighlights } from "../lib/highlighter";
import { saveRegistry } from "../lib/registry";

export function useHighlight() {
  const { registry, highlightMode, setHighlightMode, setRegistry } = useVarSyncStore();

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

  // Register the before-save guard once on mount. Read current state from the
  // store at event time so the handler doesn't stale-close over registry/mode.
  // DocumentBeforeSave is not available on all platforms — the HighlightToggle
  // component shows a persistent warning when highlights are active as a fallback.
  useEffect(() => {
    let registered = false;

    const onBeforeSave = async (_event: unknown) => {
      const { registry: r, highlightMode: on, setHighlightMode: setMode } =
        useVarSyncStore.getState();
      if (on) {
        await stripHighlights(r);
        setMode(false);
        await saveRegistry(r);
      }
    };

    try {
      Office.context.document.addHandlerAsync(
        // @ts-expect-error — not in all @types/office-js versions
        "documentBeforeSave",
        onBeforeSave,
        (result) => {
          registered = result.status === Office.AsyncResultStatus.Succeeded;
        }
      );
    } catch {
      // Not available on this platform; UI warning is the fallback
    }

    return () => {
      if (registered) {
        try {
          // @ts-expect-error
          Office.context.document.removeHandlerAsync("documentBeforeSave", {}, () => {});
        } catch {
          // noop
        }
      }
    };
  }, []); // register once — handler reads live state via getState()

  return { highlightMode, enableHighlights, disableHighlights, toggleHighlights };
}
