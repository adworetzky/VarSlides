import type { VarSyncRegistry, Binding } from "../../types";
import { REGISTRY_NAMESPACE, REGISTRY_VERSION } from "../../types";

export function emptyRegistry(): VarSyncRegistry {
  return { version: REGISTRY_VERSION, variables: [], bindings: [] };
}

function toXml(registry: VarSyncRegistry): string {
  const json = JSON.stringify(registry);
  // Escape special XML characters inside the JSON string
  const escaped = json
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return `<?xml version="1.0" encoding="utf-8"?><varsync xmlns="${REGISTRY_NAMESPACE}">${escaped}</varsync>`;
}

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

export async function saveRegistry(registry: VarSyncRegistry): Promise<void> {
  const xml = toXml(registry);

  return new Promise<void>((resolve, reject) => {
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

// Structural sanity check only — full orphan detection (missing slides/shapes) requires
// Office JS context and is done in useRegistry on load.
export function cleanOrphanedBindings(registry: VarSyncRegistry): VarSyncRegistry {
  const validVariableNames = new Set(registry.variables.map((v) => v.name));

  const cleanBindings: Binding[] = registry.bindings.filter(
    (b) =>
      validVariableNames.has(b.variableName) &&
      b.id &&
      b.shapeId &&
      b.slideIndex >= 0 &&
      b.runIndex >= 0
  );

  return { ...registry, bindings: cleanBindings };
}
