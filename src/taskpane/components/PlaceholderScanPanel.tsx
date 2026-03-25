import React, { useState } from "react";
import { scanForPlaceholders } from "../lib/finder";
import { linkWholeShape, linkInlineSelection } from "../lib/linker";
import { useRegistry } from "../hooks/useRegistry";
import type { PlaceholderMatch } from "../lib/finder";

interface Props {
  onClose: () => void;
}

interface GroupedPlaceholder {
  variableName: string;
  matches: PlaceholderMatch[];
  alreadyExists: boolean; // variable already in registry
}

export function PlaceholderScanPanel({ onClose }: Props) {
  const { registry, addVariable, addBinding } = useRegistry();
  const [scanning, setScanning] = useState(false);
  const [groups, setGroups] = useState<GroupedPlaceholder[] | null>(null);
  const [linking, setLinking] = useState<Set<string>>(new Set());
  const [linked, setLinked] = useState<Set<string>>(new Set());
  const [error, setError] = useState<string | null>(null);

  const handleScan = async () => {
    setScanning(true);
    setGroups(null);
    setError(null);
    setLinked(new Set());
    try {
      const found = await scanForPlaceholders();

      // Group by variable name
      const map = new Map<string, PlaceholderMatch[]>();
      for (const m of found) {
        if (!map.has(m.variableName)) map.set(m.variableName, []);
        map.get(m.variableName)!.push(m);
      }

      const existingNames = new Set(registry.variables.map((v) => v.name));
      const grouped: GroupedPlaceholder[] = Array.from(map.entries()).map(([name, matches]) => ({
        variableName: name,
        matches,
        alreadyExists: existingNames.has(name),
      }));

      // Sort: new variables first, then alphabetically
      grouped.sort((a, b) => {
        if (a.alreadyExists !== b.alreadyExists) return a.alreadyExists ? 1 : -1;
        return a.variableName.localeCompare(b.variableName);
      });

      setGroups(grouped);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setScanning(false);
    }
  };

  const handleLinkGroup = async (group: GroupedPlaceholder) => {
    setLinking((prev) => new Set(prev).add(group.variableName));
    try {
      // Create the variable if it doesn't exist yet (placeholder name becomes variable name + value)
      if (!group.alreadyExists) {
        await addVariable(group.variableName, group.variableName);
      }

      // After addVariable, the registry hook will have re-read; use a stable ref for linking
      const effectiveVariable = registry.variables.find((v) => v.name === group.variableName)
        ?? { name: group.variableName, value: group.variableName, color: "", createdAt: "" };

      // Link each occurrence — inline if the placeholder is a substring, whole-shape if it's the entire content
      for (const match of group.matches) {
        const existingBinding = registry.bindings.find(
          (b) => b.variableName === group.variableName && b.shapeId === match.shapeId
        );
        if (existingBinding) continue;

        const isWholeShape = match.fullParagraphText.trim() === match.rawToken;
        const binding = isWholeShape
          ? await linkWholeShape(effectiveVariable, match.shapeId, match.shapeName, match.slideIndex)
          : await linkInlineSelection(
              effectiveVariable,
              match.shapeId,
              match.slideIndex,
              match.paragraphIndex,
              match.charOffset,
              match.charOffset + match.rawToken.length,
              match.shapeName
            );
        await addBinding(binding);
      }

      setLinked((prev) => new Set(prev).add(group.variableName));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLinking((prev) => {
        const next = new Set(prev);
        next.delete(group.variableName);
        return next;
      });
    }
  };

  const handleLinkAll = async () => {
    if (!groups) return;
    for (const group of groups.filter((g) => !linked.has(g.variableName))) {
      await handleLinkGroup(group);
    }
  };

  const newGroups = groups?.filter((g) => !g.alreadyExists) ?? [];
  const existingGroups = groups?.filter((g) => g.alreadyExists) ?? [];
  const allLinked = groups !== null && groups.length > 0 && groups.every((g) => linked.has(g.variableName));

  return (
    <div className="flex flex-col gap-2 bg-neutral-800/60 border border-neutral-700 rounded p-2">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <p className="text-xs font-medium text-neutral-200">Scan for Placeholders</p>
          <p className="text-xs text-neutral-500">Find <span className="font-mono text-neutral-400">{"{{variableName}}"}</span> patterns across all slides</p>
        </div>
        <button
          onClick={onClose}
          className="text-neutral-500 hover:text-neutral-300 text-xs transition-colors ml-2 flex-shrink-0"
        >
          ×
        </button>
      </div>

      {error && (
        <div className="text-xs text-red-400 bg-red-950/40 border border-red-800/50 rounded px-2 py-1 flex items-start gap-1.5">
          <span className="flex-1">{error}</span>
          <button onClick={() => setError(null)} className="text-red-300 hover:text-red-100 flex-shrink-0">×</button>
        </div>
      )}

      {/* Scan button */}
      {groups === null && (
        <button
          onClick={() => void handleScan()}
          disabled={scanning}
          className="bg-cyan-700 hover:bg-cyan-600 disabled:bg-neutral-700 text-white text-xs rounded py-1.5 transition-colors font-medium"
        >
          {scanning ? "Scanning…" : "Scan Deck"}
        </button>
      )}

      {/* Results */}
      {groups !== null && (
        <>
          {allLinked ? (
            <div className="bg-green-950/40 border border-green-800/50 rounded px-2 py-1.5 flex items-center justify-between">
              <span className="text-xs text-green-400">All {groups.length} placeholder{groups.length !== 1 ? "s" : ""} linked</span>
              <button onClick={() => void handleScan()} className="text-xs text-neutral-400 hover:text-neutral-200">Rescan</button>
            </div>
          ) : groups.length === 0 ? (
            <p className="text-xs text-neutral-500 text-center py-2">
              No <span className="font-mono">{"{{"}</span>…<span className="font-mono">{"}}"}</span> placeholders found.
            </p>
          ) : (
            <>
              <div className="flex items-center justify-between text-xs">
                <span className="text-neutral-400">
                  {newGroups.length > 0 && `${newGroups.length} new`}
                  {newGroups.length > 0 && existingGroups.length > 0 && ", "}
                  {existingGroups.length > 0 && `${existingGroups.length} existing`}
                  {" variable"}{groups.length !== 1 ? "s" : ""}
                </span>
                <div className="flex gap-1.5">
                  {groups.some((g) => !linked.has(g.variableName)) && (
                    <button
                      onClick={() => void handleLinkAll()}
                      disabled={linking.size > 0}
                      className="text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 font-medium"
                    >
                      {linking.size > 0 ? "Linking…" : "Link All"}
                    </button>
                  )}
                  <button onClick={() => void handleScan()} className="text-neutral-400 hover:text-neutral-200">Rescan</button>
                </div>
              </div>

              <div className="flex flex-col gap-1 max-h-48 overflow-y-auto">
                {groups.map((group) => {
                  const isLinked = linked.has(group.variableName);
                  const isLinking = linking.has(group.variableName);
                  return (
                    <div
                      key={group.variableName}
                      className={`flex items-start gap-2 rounded px-2 py-1.5 text-xs ${
                        isLinked
                          ? "bg-green-950/30 border border-green-800/40"
                          : "bg-neutral-900/60 border border-neutral-700/50"
                      }`}
                    >
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-1.5">
                          <span className="font-mono text-cyan-300 font-medium">{`{{${group.variableName}}}`}</span>
                          {!group.alreadyExists && (
                            <span className="text-[10px] text-amber-400 bg-amber-950/40 px-1 rounded">new</span>
                          )}
                        </div>
                        <p className="text-neutral-500 mt-0.5">
                          {group.matches.length} occurrence{group.matches.length !== 1 ? "s" : ""} across{" "}
                          {new Set(group.matches.map((m) => m.slideIndex)).size} slide{new Set(group.matches.map((m) => m.slideIndex)).size !== 1 ? "s" : ""}
                        </p>
                      </div>
                      {isLinked ? (
                        <span className="text-green-400 flex-shrink-0">✓ linked</span>
                      ) : (
                        <button
                          onClick={() => void handleLinkGroup(group)}
                          disabled={isLinking}
                          className="text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 flex-shrink-0 transition-colors"
                        >
                          {isLinking ? "…" : group.alreadyExists ? "Link" : "Create & Link"}
                        </button>
                      )}
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}
