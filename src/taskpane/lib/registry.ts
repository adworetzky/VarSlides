/**
 * registry.ts
 *
 * Serialize / deserialize the VarSyncRegistry to/from a PowerPoint custom XML part.
 * All writes go through Zustand first; these functions are called by hooks, never
 * directly from components.
 */

import type { VarSyncRegistry, Binding } from "../../types";
import { REGISTRY_NAMESPACE, REGISTRY_VERSION } from "../../types";

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function emptyRegistry(): VarSyncRegistry {
  return { version: REGISTRY_VERSION, variables: [], bindings: [] };
}

/** Wrap JSON payload in an XML document under our namespace. */
function toXml(registry: VarSyncRegistry): string {
  const json = JSON.stringify(registry);
  // Escape special XML characters inside the JSON string
  const escaped = json
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return `<?xml version="1.0" encoding="utf-8"?><varsync xmlns="${REGISTRY_NAMESPACE}">${escaped}</varsync>`;
}

/** Extract the JSON payload from our XML part. */
function fromXml(xml: string): VarSyncRegistry {
  const match = xml.match(/<varsync[^>]*>([\s\S]*?)<\/varsync>/);
  if (!match) return emptyRegistry();
  const unescaped = match[1]
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
  try {
    return JSON.parse(unescaped) as VarSyncRegistry;
  } catch {
    return emptyRegistry();
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Load the registry from the file's custom XML part.
 * Returns an empty registry if none exists yet.
 */
export async function loadRegistry(): Promise<VarSyncRegistry> {
  return new Promise<VarSyncRegistry>((resolve, reject) => {
    Office.context.document.customXmlParts.getByNamespaceAsync(
      REGISTRY_NAMESPACE,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message ?? "Failed to query custom XML parts"));
          return;
        }

        const parts = result.value as Office.CustomXmlPart[];

        if (!parts || parts.length === 0) {
          resolve(emptyRegistry());
          return;
        }

        // Use the first matching part
        parts[0].getXmlAsync((xmlResult) => {
          if (xmlResult.status !== Office.AsyncResultStatus.Succeeded) {
            resolve(emptyRegistry());
            return;
          }
          const registry = fromXml(xmlResult.value as string);
          resolve(cleanOrphanedBindings(registry));
        });
      }
    );
  });
}

/**
 * Persist the registry to custom XML.
 * Replaces any existing part under our namespace.
 */
export async function saveRegistry(registry: VarSyncRegistry): Promise<void> {
  const xml = toXml(registry);

  return new Promise<void>((resolve, reject) => {
    // First remove any existing parts under our namespace
    Office.context.document.customXmlParts.getByNamespaceAsync(
      REGISTRY_NAMESPACE,
      (queryResult) => {
        if (queryResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(queryResult.error?.message ?? "Failed to query custom XML parts"));
          return;
        }

        const parts = queryResult.value as Office.CustomXmlPart[];
        const deletionPromises = (parts ?? []).map(
          (part) =>
            new Promise<void>((res, rej) => {
              part.deleteAsync((deleteResult) => {
                if (deleteResult.status !== Office.AsyncResultStatus.Succeeded) {
                  rej(new Error(deleteResult.error?.message ?? "Failed to delete custom XML part"));
                } else {
                  res();
                }
              });
            })
        );

        Promise.all(deletionPromises)
          .then(() => {
            Office.context.document.customXmlParts.addAsync(xml, (addResult) => {
              if (addResult.status !== Office.AsyncResultStatus.Succeeded) {
                reject(new Error(addResult.error?.message ?? "Failed to add custom XML part"));
              } else {
                resolve();
              }
            });
          })
          .catch(reject);
      }
    );
  });
}

/**
 * Remove bindings that reference shapes or slides that no longer exist.
 * Called automatically on every load before data surfaces to the store.
 *
 * Note: Without full slide/shape enumeration at load time this performs a
 * structural sanity check only (removes obviously malformed records).
 * Full orphan detection requires Office JS context and is done in useRegistry.
 */
export function cleanOrphanedBindings(registry: VarSyncRegistry): VarSyncRegistry {
  const validVariableNames = new Set(registry.variables.map((v) => v.name));

  const cleanBindings: Binding[] = registry.bindings.filter((b) => {
    // Drop bindings for variables that no longer exist
    if (!validVariableNames.has(b.variableName)) return false;
    // Drop structurally invalid bindings
    if (!b.id || !b.shapeId || b.slideIndex < 0 || b.runIndex < 0) return false;
    return true;
  });

  return { ...registry, bindings: cleanBindings };
}
