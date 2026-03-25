import { isLinkableShapeType } from "./linker";

/** Regex for `{{variableName}}` placeholders — captures the name inside the braces */
const PLACEHOLDER_RE = /\{\{([^}]+)\}\}/g;

export interface PlaceholderMatch {
  /** Variable name extracted from `{{name}}` */
  variableName: string;
  /** The full token as it appears in the slide, e.g. `{{AUM}}` */
  rawToken: string;
  shapeId: string;
  shapeName: string;
  slideIndex: number;
  paragraphIndex: number;
  charOffset: number;
  fullParagraphText: string;
}

/**
 * Scan every linkable shape in the deck for `{{variableName}}` placeholders.
 * Uses the same 4-pass batching strategy as findTextInDeck — no N+1 syncs.
 */
export async function scanForPlaceholders(): Promise<PlaceholderMatch[]> {
  const matches: PlaceholderMatch[] = [];

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (const slide of slides.items) {
      slide.shapes.load("items");
    }
    await context.sync();

    const targets: { shape: PowerPoint.Shape; slideIndex: number }[] = [];
    for (let si = 0; si < slides.items.length; si++) {
      for (const shape of slides.items[si].shapes.items) {
        shape.load("id, name, type");
        targets.push({ shape, slideIndex: si });
      }
    }
    await context.sync();

    const linkable = targets.filter((t) => {
      try { return isLinkableShapeType(t.shape.type as string); } catch { return false; }
    });

    for (const { shape } of linkable) {
      try { shape.textFrame.textRange.paragraphs.load("items"); } catch { /* no text frame */ }
    }
    await context.sync();

    for (const { shape } of linkable) {
      try {
        for (const para of shape.textFrame.textRange.paragraphs.items) {
          para.textRange.load("text");
        }
      } catch { /* skip */ }
    }
    await context.sync();

    for (const { shape, slideIndex } of linkable) {
      try {
        const paras = shape.textFrame.textRange.paragraphs.items;
        for (let pi = 0; pi < paras.length; pi++) {
          const text = paras[pi].textRange.text;
          PLACEHOLDER_RE.lastIndex = 0;
          let m: RegExpExecArray | null;
          while ((m = PLACEHOLDER_RE.exec(text)) !== null) {
            matches.push({
              variableName: m[1].trim(),
              rawToken: m[0],
              shapeId: shape.id,
              shapeName: shape.name,
              slideIndex,
              paragraphIndex: pi,
              charOffset: m.index,
              fullParagraphText: text,
            });
          }
        }
      } catch { /* skip */ }
    }
  });

  return matches;
}

export interface TextMatch {
  shapeId: string;
  shapeName: string;
  slideIndex: number;
  paragraphIndex: number;
  charOffset: number;
  /** Full paragraph text — used to show context around the match in the UI */
  fullParagraphText: string;
  /** True if this shapeId already has a binding for this variable */
  alreadyLinked: boolean;
}

/**
 * Scan every linkable shape in the deck for paragraphs containing `searchText`.
 * Uses 4 fixed context.sync() calls regardless of deck size — no N+1.
 *
 * `linkedShapeIds` — set of shapeIds that already have a binding for this
 * variable, used to flag duplicates in the results.
 */
export async function findTextInDeck(
  searchText: string,
  linkedShapeIds: Set<string> = new Set()
): Promise<TextMatch[]> {
  if (!searchText) return [];

  const matches: TextMatch[] = [];

  await PowerPoint.run(async (context) => {
    // Pass 1: load slides
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    // Pass 2: load all shape collections
    for (const slide of slides.items) {
      slide.shapes.load("items");
    }
    await context.sync();

    // Pass 3: load shape metadata (id, name, type)
    const targets: { shape: PowerPoint.Shape; slideIndex: number }[] = [];
    for (let si = 0; si < slides.items.length; si++) {
      for (const shape of slides.items[si].shapes.items) {
        shape.load("id, name, type");
        targets.push({ shape, slideIndex: si });
      }
    }
    await context.sync();

    // Filter to linkable shapes and batch-load their paragraph collections
    const linkable = targets.filter((t) => {
      try {
        return isLinkableShapeType(t.shape.type as string);
      } catch {
        return false;
      }
    });

    for (const { shape } of linkable) {
      try {
        shape.textFrame.textRange.paragraphs.load("items");
      } catch {
        // Shape has no text frame — skip
      }
    }
    await context.sync();

    // Batch-load all paragraph texts
    for (const { shape } of linkable) {
      try {
        for (const para of shape.textFrame.textRange.paragraphs.items) {
          para.textRange.load("text");
        }
      } catch {
        // Skip shapes whose paragraph collection failed to load
      }
    }
    await context.sync();

    // Collect matches
    for (const { shape, slideIndex } of linkable) {
      try {
        const paras = shape.textFrame.textRange.paragraphs.items;
        for (let pi = 0; pi < paras.length; pi++) {
          const text = paras[pi].textRange.text;
          if (!text.includes(searchText)) continue;

          let charOffset = 0;
          while (true) {
            const idx = text.indexOf(searchText, charOffset);
            if (idx === -1) break;
            matches.push({
              shapeId: shape.id,
              shapeName: shape.name,
              slideIndex,
              paragraphIndex: pi,
              charOffset: idx,
              fullParagraphText: text,
              alreadyLinked: linkedShapeIds.has(shape.id),
            });
            charOffset = idx + searchText.length;
          }
        }
      } catch {
        // Skip shapes that failed during text access
      }
    }
  });

  return matches;
}

/**
 * Pure helper — extract a display snippet around a match for the UI.
 * Returns up to `context` characters on each side of the match, with
 * the matched portion indicated by its position.
 */
export function matchSnippet(
  fullText: string,
  matchText: string,
  charOffset: number,
  context = 20
): { before: string; match: string; after: string } {
  const start = Math.max(0, charOffset - context);
  const end = Math.min(fullText.length, charOffset + matchText.length + context);
  return {
    before: (start > 0 ? "…" : "") + fullText.slice(start, charOffset),
    match: matchText,
    after: fullText.slice(charOffset + matchText.length, end) + (end < fullText.length ? "…" : ""),
  };
}
