import React, { useState } from "react";
import { findTextInDeck, matchSnippet } from "../lib/finder";
import { linkWholeShape, linkInlineSelection } from "../lib/linker";
import { useRegistry } from "../hooks/useRegistry";
import type { Variable } from "../../types";
import type { TextMatch } from "../lib/finder";

interface Props {
  variable: Variable;
  onClose: () => void;
}

export function FindLinkPanel({ variable, onClose }: Props) {
  const { registry, addBinding } = useRegistry();
  const [scanning, setScanning] = useState(false);
  const [matches, setMatches] = useState<TextMatch[] | null>(null);
  const [linking, setLinking] = useState<Set<string>>(new Set());
  const [linked, setLinked] = useState<Set<string>>(new Set());
  const [linkProgress, setLinkProgress] = useState<{ current: number; total: number } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const linkedShapeIds = new Set(
    registry.bindings
      .filter((b) => b.variableName === variable.name)
      .map((b) => b.shapeId)
  );

  const handleScan = async () => {
    setScanning(true);
    setMatches(null);
    setError(null);
    setLinked(new Set());
    try {
      const results = await findTextInDeck(variable.value, linkedShapeIds);
      setMatches(results);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setScanning(false);
    }
  };

  const matchKey = (m: TextMatch) =>
    `${m.shapeId}:${m.slideIndex}:${m.paragraphIndex}:${m.charOffset}`;

  const handleLinkMatch = async (match: TextMatch) => {
    const key = matchKey(match);
    setLinking((prev) => new Set(prev).add(key));
    try {
      const isWholeShape = match.fullParagraphText.trim() === variable.value;
      const binding = isWholeShape
        ? await linkWholeShape(variable, match.shapeId, match.shapeName, match.slideIndex)
        : await linkInlineSelection(
            variable,
            match.shapeId,
            match.slideIndex,
            match.paragraphIndex,
            match.charOffset,
            match.charOffset + variable.value.length,
            match.shapeName
          );
      await addBinding(binding);
      setLinked((prev) => new Set(prev).add(key));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLinking((prev) => {
        const next = new Set(prev);
        next.delete(key);
        return next;
      });
    }
  };

  const handleLinkAll = async () => {
    if (!matches) return;
    const toLink = matches.filter((m) => !m.alreadyLinked && !linked.has(matchKey(m)));
    setLinkProgress({ current: 0, total: toLink.length });
    for (let i = 0; i < toLink.length; i++) {
      setLinkProgress({ current: i + 1, total: toLink.length });
      await handleLinkMatch(toLink[i]);
    }
    setLinkProgress(null);
  };

  const newMatches = matches?.filter((m) => !m.alreadyLinked) ?? [];
  const alreadyLinkedMatches = matches?.filter((m) => m.alreadyLinked) ?? [];
  const allDone = matches !== null && newMatches.length > 0 && linked.size >= newMatches.length;

  return (
    <div className="flex flex-col gap-2 bg-neutral-800/60 border border-neutral-700 rounded p-2">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <p className="text-xs font-medium text-neutral-200">
            Find &amp; Link:{" "}
            <span className="font-mono text-cyan-300">{variable.value}</span>
          </p>
          <p className="text-xs text-neutral-500">
            Scan all slides for text matching this value
          </p>
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
      {matches === null && (
        <button
          onClick={() => void handleScan()}
          disabled={scanning}
          className="bg-cyan-700 hover:bg-cyan-600 disabled:bg-neutral-700 text-white text-xs rounded py-1.5 transition-colors font-medium"
        >
          {scanning ? "Scanning…" : "Scan Deck"}
        </button>
      )}

      {/* Results */}
      {matches !== null && (
        <>
          {/* All done banner */}
          {allDone ? (
            <div className="bg-green-950/40 border border-green-800/50 rounded px-2 py-1.5 flex items-center justify-between">
              <span className="text-xs text-green-400">All {newMatches.length} occurrence{newMatches.length !== 1 ? "s" : ""} linked</span>
              <button
                onClick={() => void handleScan()}
                className="text-xs text-neutral-400 hover:text-neutral-200 transition-colors"
              >
                Rescan
              </button>
            </div>
          ) : (
            <div className="flex items-center justify-between text-xs">
              <span className="text-neutral-400">
                {newMatches.length} new match{newMatches.length !== 1 ? "es" : ""}
                {alreadyLinkedMatches.length > 0 &&
                  `, ${alreadyLinkedMatches.length} already linked`}
              </span>
              <div className="flex gap-1.5">
                {newMatches.length > 0 && (
                  <button
                    onClick={() => void handleLinkAll()}
                    disabled={linking.size > 0}
                    className="text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 font-medium transition-colors"
                  >
                    {linkProgress
                      ? `Linking ${linkProgress.current}/${linkProgress.total}…`
                      : linking.size > 0 ? "Linking…" : "Link All"}
                  </button>
                )}
                <button
                  onClick={() => void handleScan()}
                  disabled={scanning}
                  className="text-neutral-400 hover:text-neutral-200 transition-colors"
                >
                  Rescan
                </button>
              </div>
            </div>
          )}

          {matches.length === 0 && (
            <p className="text-xs text-neutral-500 text-center py-2">
              No occurrences of{" "}
              <span className="font-mono text-neutral-300">{variable.value}</span>{" "}
              found in the deck.
            </p>
          )}

          <div className="flex flex-col gap-1 max-h-48 overflow-y-auto">
            {matches.map((m) => {
              const key = matchKey(m);
              const isLinked = linked.has(key) || m.alreadyLinked;
              const isLinking = linking.has(key);
              const snippet = matchSnippet(m.fullParagraphText, variable.value, m.charOffset);

              return (
                <div
                  key={key}
                  className={`flex items-start gap-2 rounded px-2 py-1 text-xs ${
                    isLinked
                      ? "bg-green-950/30 border border-green-800/40"
                      : "bg-neutral-900/60 border border-neutral-700/50"
                  }`}
                >
                  <div className="flex-1 min-w-0">
                    <span className="text-neutral-400">
                      Slide {m.slideIndex + 1} · {m.shapeName}
                    </span>
                    <p className="text-neutral-300 truncate mt-0.5">
                      <span className="text-neutral-500">{snippet.before}</span>
                      <span className="text-cyan-300 font-medium">{snippet.match}</span>
                      <span className="text-neutral-500">{snippet.after}</span>
                    </p>
                  </div>
                  {isLinked ? (
                    <span className="text-green-400 flex-shrink-0 pt-0.5">
                      {m.alreadyLinked ? "was linked" : "✓ linked"}
                    </span>
                  ) : (
                    <button
                      onClick={() => void handleLinkMatch(m)}
                      disabled={isLinking}
                      className="text-cyan-400 hover:text-cyan-300 disabled:text-neutral-600 flex-shrink-0 pt-0.5 transition-colors"
                    >
                      {isLinking ? "…" : "Link"}
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}
