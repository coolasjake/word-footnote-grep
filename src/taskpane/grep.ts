export type NoteKind = "footnote" | "endnote" | "both";

export interface GrepFlags {
  global: boolean;
  ignoreCase: boolean;
  multiline: boolean;
  dotAll: boolean;
}

export interface GrepOptions {
  pattern: string;
  replacement: string;
  flags: GrepFlags;
  noteKind: NoteKind;
}

export interface GrepMatch {
  noteIndex: number;
  noteKind: "footnote" | "endnote";
  originalText: string;
  newText: string;
  referenceText?: string;
  matchCount: number;
}

export interface GrepResult {
  matches: GrepMatch[];
  totalMatchCount: number;
  error?: string;
}

export function buildRegExp(pattern: string, flags: GrepFlags): RegExp {
  let flagStr = "";
  if (flags.global) flagStr += "g";
  if (flags.ignoreCase) flagStr += "i";
  if (flags.multiline) flagStr += "m";
  if (flags.dotAll) flagStr += "s";

  return new RegExp(pattern, flagStr);
}

export function applyGrep(text: string, pattern: RegExp, replacement: string): { newText: string; matchCount: number } {
  const testPattern = new RegExp(
    pattern.source,
    pattern.flags.includes("g") ? pattern.flags : pattern.flags + "g"
  );

  const matches = text.match(testPattern);
  const matchCount = matches ? matches.length : 0;

  if (matchCount === 0) {
    return { newText: text, matchCount: 0 };
  }

  const newText = text.replace(pattern, replacement);
  return { newText, matchCount };
}

export function previewGrep(notes: Array<{ index: number; kind: "footnote" | "endnote"; text: string; referenceText?: string }>, options: GrepOptions): GrepResult {
  try {
    const pattern = buildRegExp(options.pattern, options.flags);
    const matches: GrepMatch[] = [];
    let totalMatchCount = 0;

    for (const note of notes) {
      const { newText, matchCount } = applyGrep(note.text, pattern, options.replacement);
      if (matchCount > 0) {
        matches.push({
          noteIndex: note.index,
          noteKind: note.kind,
          originalText: note.text,
          newText,
          referenceText: note.referenceText,
          matchCount,
        });
        totalMatchCount += matchCount;
      }
    }

    return { matches, totalMatchCount };
  } catch (err) {
    return {
      matches: [],
      totalMatchCount: 0,
      error: err instanceof Error ? err.message : String(err),
    };
  }
}

export function countMatches(text: string, pattern: RegExp): number {
  const globalPattern = new RegExp(
    pattern.source,
    pattern.flags.includes("g") ? pattern.flags : pattern.flags + "g"
  );
  return text.match(globalPattern)?.length ?? 0;
}
