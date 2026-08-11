import {
  GrepMatch,
  GrepOptions,
  NoteKind,
  previewGrep,
} from "./grep";


export interface LoadedNote {
  index: number;
  kind: "footnote" | "endnote";
  text: string;
}


export interface ItalicisedComma {
  noteIndex: number;
  occurrence: number;
  contextBefore: string;
  contextAfter: string;
}


export interface SourceReference {
  noteIndex: number;
  sourceIndex: number;
  reference: string;
  source: string;
  normalizedSource: string;
}


export interface SourceGroup {
  source: string;
  normalizedSource: string;
  references: string[];
}


/* --------------------------------------------------------------------------
   GENERAL NOTE LOADING
-------------------------------------------------------------------------- */

async function loadNotes(
  context: Word.RequestContext,
  kind: NoteKind
): Promise<LoadedNote[]> {
  const notes: LoadedNote[] = [];

  if (kind === "footnote" || kind === "both") {
    const footnotes = context.document.body.footnotes;

    footnotes.load("items");
    await context.sync();

    const rangeObjects: Word.Range[] = [];

    for (let i = 0; i < footnotes.items.length; i++) {
      const range = footnotes.items[i].body.getRange();

      range.load("text");
      rangeObjects.push(range);
    }

    await context.sync();

    for (let i = 0; i < footnotes.items.length; i++) {
      notes.push({
        index: i + 1,
        kind: "footnote",
        text: rangeObjects[i].text,
      });
    }
  }


  if (kind === "endnote" || kind === "both") {
    const endnotes = context.document.body.endnotes;

    endnotes.load("items");
    await context.sync();

    const rangeObjects: Word.Range[] = [];

    for (let i = 0; i < endnotes.items.length; i++) {
      const range = endnotes.items[i].body.getRange();

      range.load("text");
      rangeObjects.push(range);
    }

    await context.sync();

    for (let i = 0; i < endnotes.items.length; i++) {
      notes.push({
        index: i + 1,
        kind: "endnote",
        text: rangeObjects[i].text,
      });
    }
  }

  return notes;
}


export async function loadAllNotes(
  kind: NoteKind
): Promise<LoadedNote[]> {
  return Word.run(async (context) =>
    loadNotes(context, kind)
  );
}


/* --------------------------------------------------------------------------
   EXISTING GREP FUNCTIONALITY
-------------------------------------------------------------------------- */

export async function previewFootnoteGrep(
  options: GrepOptions
) {
  const notes = await loadAllNotes(options.noteKind);

  return previewGrep(notes, options);
}


async function applyMatch(
  context: Word.RequestContext,
  match: GrepMatch
): Promise<void> {
  const collection =
    match.noteKind === "footnote"
      ? context.document.body.footnotes
      : context.document.body.endnotes;

  collection.load("items");
  await context.sync();

  const note = collection.items[match.noteIndex - 1];
  const range = note.body.getRange();

  range.insertText(
    match.newText,
    Word.InsertLocation.replace
  );
}


export async function applyGrepMatches(
  matches: GrepMatch[]
): Promise<number> {
  let applied = 0;

  await Word.run(async (context) => {
    const footnotes = context.document.body.footnotes;
    const endnotes = context.document.body.endnotes;

    footnotes.load("items");
    endnotes.load("items");

    await context.sync();

    for (const match of matches) {
      const collection =
        match.noteKind === "footnote"
          ? footnotes
          : endnotes;

      const note =
        collection.items[match.noteIndex - 1];

      const range = note.body.getRange();

      range.insertText(
        match.newText,
        Word.InsertLocation.replace
      );

      applied++;
    }

    await context.sync();
  });

  return applied;
}


export async function applyGrepOptions(
  options: GrepOptions
): Promise<{
  applied: number;
  result: ReturnType<typeof previewGrep>;
}> {
  const result = await previewFootnoteGrep(options);

  if (
    result.error ||
    result.matches.length === 0
  ) {
    return {
      applied: 0,
      result,
    };
  }

  const applied = await applyGrepMatches(
    result.matches
  );

  return {
    applied,
    result,
  };
}


/* --------------------------------------------------------------------------
   NOTE COUNTS
-------------------------------------------------------------------------- */

export async function getNoteCounts(): Promise<{
  footnotes: number;
  endnotes: number;
}> {
  return Word.run(async (context) => {
    const footnotes =
      context.document.body.footnotes;

    const endnotes =
      context.document.body.endnotes;

    footnotes.load("items");
    endnotes.load("items");

    await context.sync();

    return {
      footnotes: footnotes.items.length,
      endnotes: endnotes.items.length,
    };
  });
}


/* --------------------------------------------------------------------------
   ITALICISED COMMAS
-------------------------------------------------------------------------- */

function makeContext(
  text: string,
  commaPosition: number
): {
  before: string;
  after: string;
} {
  const contextLength = 45;

  const beforeStart = Math.max(
    0,
    commaPosition - contextLength
  );

  const afterEnd = Math.min(
    text.length,
    commaPosition + contextLength + 1
  );

  return {
    before: text.slice(
      beforeStart,
      commaPosition
    ),

    after: text.slice(
      commaPosition + 1,
      afterEnd
    ),
  };
}


export async function findItalicisedCommas(): Promise<
  ItalicisedComma[]
> {
  return Word.run(async (context) => {
    const footnotes =
      context.document.body.footnotes;

    footnotes.load("items");
    await context.sync();

    const searchResults: {
      noteIndex: number;
      ranges: Word.RangeCollection;
      bodyText: Word.Range;
    }[] = [];

    for (let i = 0; i < footnotes.items.length; i++) {
      const body =
        footnotes.items[i].body;

      const bodyRange = body.getRange();

      bodyRange.load("text");

      const ranges = body.search(",", {
        matchCase: true,
      });

      ranges.load("items");

      searchResults.push({
        noteIndex: i + 1,
        ranges,
        bodyText: bodyRange,
      });
    }

    await context.sync();

    const result: ItalicisedComma[] = [];

    for (const item of searchResults) {
      const ranges = item.ranges;

      /*
       * Load the formatting of every comma found in this
       * footnote. Font.italic is used rather than replacing
       * the text, so the correction affects only formatting.
       */
      for (const range of ranges.items) {
        range.load("text");
        range.font.load("italic");
      }
    }

    await context.sync();

    for (const item of searchResults) {
      const fullText = item.bodyText.text;
      let occurrence = 0;

      for (const range of item.ranges.items) {
        if (!range.font.italic) {
          continue;
        }

        occurrence++;

        /*
         * Find the actual comma position in the original
         * footnote text. Using indexOf from the running
         * position handles repeated commas correctly.
         */
        const previousOccurrences =
          item.ranges.items
            .slice(
              0,
              item.ranges.items.indexOf(range)
            )
            .length;

        let commaPosition = -1;
        let searchFrom = 0;

        for (
          let i = 0;
          i <= previousOccurrences;
          i++
        ) {
          commaPosition = fullText.indexOf(
            ",",
            searchFrom
          );

          if (commaPosition === -1) {
            break;
          }

          searchFrom = commaPosition + 1;
        }

        if (commaPosition === -1) {
          continue;
        }

        const contextText =
          makeContext(
            fullText,
            commaPosition
          );

        result.push({
          noteIndex: item.noteIndex,
          occurrence,
          contextBefore: contextText.before,
          contextAfter: contextText.after,
        });
      }
    }

    return result;
  });
}


export async function fixItalicisedCommas(): Promise<number> {
  return Word.run(async (context) => {
    const footnotes =
      context.document.body.footnotes;

    footnotes.load("items");
    await context.sync();

    const allRanges: Word.Range[] = [];

    for (let i = 0; i < footnotes.items.length; i++) {
      const ranges =
        footnotes.items[i].body.search(",", {
          matchCase: true,
        });

      ranges.load("items");

      /*
       * Keep the comma ranges so that we can inspect their
       * formatting after the batch sync.
       */
      allRanges.push(...([] as Word.Range[]));

      (
        allRanges as unknown as {
          push: (
            ...items: Word.Range[]
          ) => number;
        }
      );
    }

    /*
     * Rebuild the ranges in a structure that preserves the
     * footnote boundaries.
     */
    const footnoteRanges: Word.RangeCollection[] = [];

    for (let i = 0; i < footnotes.items.length; i++) {
      const ranges =
        footnotes.items[i].body.search(",", {
          matchCase: true,
        });

      ranges.load("items");
      footnoteRanges.push(ranges);
    }

    await context.sync();

    let fixed = 0;

    for (const ranges of footnoteRanges) {
      for (const range of ranges.items) {
        range.font.load("italic");
      }
    }

    await context.sync();

    for (const ranges of footnoteRanges) {
      for (const range of ranges.items) {
        if (range.font.italic) {
          range.font.italic = false;
          fixed++;
        }
      }
    }

    await context.sync();

    return fixed;
  });
}


/* --------------------------------------------------------------------------
   SOURCE EXTRACTION
-------------------------------------------------------------------------- */

/**
 * Normalises source text for grouping.
 *
 * This deliberately uses conservative normalisation:
 *
 *   "Harry  2021"
 *   " Harry 2021 "
 *
 * are treated as the same source, while substantially
 * different citations are not automatically merged.
 */
function normalizeSource(source: string): string {
  return source
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(:])\s+/g, "$1")
    .toLocaleLowerCase();
}


export async function getFootnoteSources(): Promise<
  SourceReference[]
> {
  const notes = await loadAllNotes("footnote");

  const references: SourceReference[] = [];

  for (const note of notes) {
    /*
     * Each semicolon represents a separate source.
     *
     * Empty pieces are ignored so that accidental
     * double-semicolons don't produce empty entries.
     */
    const sources = note.text
      .split(";")
      .map((source) => source.trim())
      .filter((source) => source.length > 0);

    sources.forEach((source, index) => {
      references.push({
        noteIndex: note.index,
        sourceIndex: index + 1,
        reference: `${note.index}(${index + 1})`,
        source,
        normalizedSource:
          normalizeSource(source),
      });
    });
  }

  return references;
}