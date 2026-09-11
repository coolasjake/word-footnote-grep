// Imports the shared grep types and preview helper used by this module.
import {
  GrepMatch,
  GrepOptions,
  NoteKind,
  previewGrep,
} from "./grep";


// Represents a loaded footnote/endnote and the metadata needed by the UI and source-processing logic.
export interface LoadedNote {
  index: number;
  kind: "footnote" | "endnote";
  text: string;
  referenceText: string;
  isCustomReference: boolean;
  customReferenceText?: string;
}


// Describes an italicised comma found inside a footnote, including enough surrounding text to identify it.
export interface ItalicisedComma {
  noteIndex: number;
  occurrence: number;
  contextBefore: string;
  contextAfter: string;
}


// Represents one extracted source citation and the metadata used to link, group, and display it.
export interface SourceReference {
  noteIndex: number;
  sourceIndex: number;
  occurrence: number;
  reference: string;
  source: string;
  fullNoteText: string;
  sourceCount: number;
  directReferenceTarget?: string;
  directReferenceTargetIndex?: number;
  directReferencePrefix?: string;
  shortNameDeclaration?: string;
  isIbid?: boolean;
  ibidTargetNoteIndex?: number;
  normalizedSource: string;
}


// Represents a group of source references that normalise to the same source text.
export interface SourceGroup {
  source: string;
  normalizedSource: string;
  references: SourceReference[];
  problemReferences?: SourceReference[];
  warning?: string;
  error?: string;
}


// Citation phrases are ignored when source text is compared for likely matches.
export const keywordPattern =
  /\b(?:quoting|quoted in|citing|cited in|discussing|discussed in|and|&|see|see, eg|see also|see especially|see generally|cf|but see|eg)\b/gi;


/* --------------------------------------------------------------------------
   GENERAL NOTE LOADING
-------------------------------------------------------------------------- */

// Loads footnotes and/or endnotes from Word, optionally including reference and custom-reference metadata.
async function loadNotes(
  context: Word.RequestContext,
  kind: NoteKind,
  includeReferenceMetadata = false
): Promise<LoadedNote[]> {
  const notes: LoadedNote[] = [];

  if (kind === "footnote" || kind === "both") {
    const footnotes = context.document.body.footnotes;

    footnotes.load("items");
    await context.sync();

    const rangeObjects: Word.Range[] = [];
    const referenceObjects: Word.Range[] = [];
    const bodyObjects: Word.Body[] = [];

    for (let i = 0; i < footnotes.items.length; i++) {
      const range = footnotes.items[i].body.getRange();
      const reference = footnotes.items[i].reference;

      range.load("text");
      if (includeReferenceMetadata) {
        reference.load("text");
      }
      rangeObjects.push(range);
      if (includeReferenceMetadata) {
        referenceObjects.push(reference);
        bodyObjects.push(footnotes.items[i].body);
      }
    }

    await context.sync();

    const bodyOoxmlObjects = includeReferenceMetadata
      ? bodyObjects.map((body, index) =>
          isAutomaticReference(referenceObjects[index].text)
            ? undefined
            : body.getOoxml()
        )
      : [];

    if (includeReferenceMetadata) {
      await context.sync();
    }

    for (let i = 0; i < footnotes.items.length; i++) {
      const bodyOoxml = bodyOoxmlObjects[i];
      const customReferenceText = bodyOoxml
        ? readCustomReference(bodyOoxml.value)
        : undefined;

      notes.push({
        index: i + 1,
        kind: "footnote",
        text: rangeObjects[i].text,
        referenceText: includeReferenceMetadata
          ? referenceObjects[i].text
          : "",
        isCustomReference: Boolean(customReferenceText),
        customReferenceText,
      });
    }
  }


  if (kind === "endnote" || kind === "both") {
    const endnotes = context.document.body.endnotes;

    endnotes.load("items");
    await context.sync();

    const rangeObjects: Word.Range[] = [];
    const referenceObjects: Word.Range[] = [];
    const bodyObjects: Word.Body[] = [];

    for (let i = 0; i < endnotes.items.length; i++) {
      const range = endnotes.items[i].body.getRange();
      const reference = endnotes.items[i].reference;

      range.load("text");
      if (includeReferenceMetadata) {
        reference.load("text");
      }
      rangeObjects.push(range);
      if (includeReferenceMetadata) {
        referenceObjects.push(reference);
        bodyObjects.push(endnotes.items[i].body);
      }
    }

    await context.sync();

    const bodyOoxmlObjects = includeReferenceMetadata
      ? bodyObjects.map((body, index) =>
          isAutomaticReference(referenceObjects[index].text)
            ? undefined
            : body.getOoxml()
        )
      : [];

    if (includeReferenceMetadata) {
      await context.sync();
    }

    for (let i = 0; i < endnotes.items.length; i++) {
      const bodyOoxml = bodyOoxmlObjects[i];
      const customReferenceText = bodyOoxml
        ? readCustomReference(bodyOoxml.value)
        : undefined;

      notes.push({
        index: i + 1,
        kind: "endnote",
        text: rangeObjects[i].text,
        referenceText: includeReferenceMetadata
          ? referenceObjects[i].text
          : "",
        isCustomReference: Boolean(customReferenceText),
        customReferenceText,
      });
    }
  }

  return notes;
}


// Public wrapper that loads all requested notes through Word.run and returns plain LoadedNote objects.
export async function loadAllNotes(
  kind: NoteKind,
  includeReferenceMetadata = false
): Promise<LoadedNote[]> {
  return Word.run(async (context) =>
    loadNotes(context, kind, includeReferenceMetadata)
  );
}


// Selects a specific occurrence of text inside a footnote or endnote and moves Word's selection to it.
export async function navigateToNote(
  kind: "footnote" | "endnote",
  noteIndex: number,
  searchText: string,
  occurrence = 1
): Promise<void> {
  const navigationText = searchText.length > 180
    ? searchText.slice(0, 180).replace(/\s+\S*$/, "")
    : searchText;

  await Word.run(async (context) => {
    const collection = kind === "footnote"
      ? context.document.body.footnotes
      : context.document.body.endnotes;

    collection.load("items");
    await context.sync();

    const note = collection.items[noteIndex - 1];

    if (!note) {
      throw new Error(`Could not find ${kind} ${noteIndex}.`);
    }

    const matches = note.body.search(navigationText, {
      matchCase: true,
      matchWholeWord: false,
    });
    matches.load("items");
    await context.sync();

    const target = matches.items[occurrence - 1];

    if (!target) {
      throw new Error(
        `Could not find the selected text in ${kind} ${noteIndex}.`
      );
    }

    target.select();
    await context.sync();
  });
}


/* --------------------------------------------------------------------------
   EXISTING GREP FUNCTIONALITY
-------------------------------------------------------------------------- */

// Runs the existing grep preview logic against the requested set of notes without changing the document.
export async function previewFootnoteGrep(
  options: GrepOptions
) {
  const notes = await loadAllNotes(options.noteKind);

  return previewGrep(notes, options);
}


// Applies one grep match by replacing the complete body text of the matching note.
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


// Applies a collection of grep matches in one Word.run operation and returns the number applied.
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


// Previews grep changes first, then applies them only when there is no error and at least one match.
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

// Returns the current number of footnotes and endnotes in the document.
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

// Builds a short piece of surrounding text for an identified comma, for display in the results list.
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


// Finds commas in footnotes that are currently italicised and returns their locations and context.
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
          occurrence:
            item.ranges.items.indexOf(range) + 1,
          contextBefore: contextText.before,
          contextAfter: contextText.after,
        });
      }
    }

    return result;
  });
}


// Removes italic formatting from every italicised comma in the document and returns the number fixed.
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
// Normalises source text conservatively so equivalent spacing/punctuation variants can be grouped.
function normalizeSource(source: string): string {
  return source
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(:])\s+/g, "$1")
    .toLocaleLowerCase();
}


// Removes the leading marker characters Word stores in note text before source extraction.
function removeLeadingNoteMarkers(
  text: string,
  referenceText: string
): string {
  if (isAutomaticReference(referenceText)) {
    return text.replace(/^[\u0000-\u001f\u007f]+/, "");
  }

  return text.replace(/^\(\s*/, "");
}


// Reads a custom footnote reference symbol from the note's OOXML, when one is present.
function readCustomReference(xml: string): string | undefined {
  const marker = xml.match(
    /<w:footnoteRef\b[^>]*w:customMarkFollows="1"[^>]*\/?\s*>/
  );

  if (!marker) {
    return undefined;
  }

  const followingXml = xml.slice(
    (marker.index ?? 0) + marker[0].length
  );
  const symbol = followingXml.match(
    /<w:sym\b[^>]*w:char="([0-9A-Fa-f]{4,6})"[^>]*\/?\s*>/
  );

  if (!symbol) {
    return "*";
  }

  return String.fromCodePoint(parseInt(symbol[1], 16));
}


// Determines whether Word's reference text represents an automatically numbered note.
function isAutomaticReference(referenceText: string): boolean {
  return referenceText.trimStart().charCodeAt(0) === 0x0002;
}


// Splits a note's text into separate source citations using semicolons while avoiding common numeric cases.
function splitSources(text: string): string[] {
  return text
    .split(/;(?=\s*(?:[^\d\s]|$))/)
    .map((source) => source.trim())
    .filter((source) => source.length > 0);
}


// Detects a direct reference such as "(n 12)" and extracts its target note number and descriptive prefix.
function readDirectReference(source: string): {
  target: string;
  prefix: string;
} | undefined {
  const cleanSource = source
    .replace(/^[\u0000-\u001f\u007f]+|[\u0000-\u001f\u007f]+$/g, "")
    .trim();
  const match = cleanSource.match(/\(n (\d+)\)/);

  if (!match || match.index === undefined) {
    return undefined;
  }

  const precedingText = cleanSource.slice(0, match.index);
  const punctuationIndex = Math.max(
    precedingText.lastIndexOf(","),
    precedingText.lastIndexOf("."),
    precedingText.lastIndexOf(":"),
    precedingText.lastIndexOf(";"),
    precedingText.lastIndexOf("!"),
    precedingText.lastIndexOf("?")
  );
  let boundary = punctuationIndex + 1;
  let keywordMatch: RegExpExecArray | null;

  while ((keywordMatch = keywordPattern.exec(precedingText)) !== null) {
    boundary = Math.max(
      boundary,
      keywordMatch.index + keywordMatch[0].length
    );
  }

  const words = precedingText
    .slice(boundary)
    .trim()
    .split(/\s+/)
    .filter((word) => word.length > 0);
  const prefix = words.slice(-3).join(" ");

  return {
    target: match[1],
    prefix,
  };
}


// Detects a short-name declaration at the end of a source citation and returns the declared short name.
function readShortNameDeclaration(source: string): string | undefined {
  const match = source.match(
    /\(\s*['"\u2018\u201C]\s*(.*?)\s*['"\u2019\u201D]\s*\)\s*[.!?;:]?$/
  );

  return match?.[1].trim() || undefined;
}


// Checks whether a source citation begins with an Ibid marker.
function isIbidSource(source: string): boolean {
  return /^ibid\b/i.test(source.trim());
}


// Extracts all source references from the document and resolves their citation metadata.
export async function getFootnoteSources(): Promise<
  SourceReference[]
> {
  const notes = await loadAllNotes("footnote", true);
  const references: SourceReference[] = [];
  let automaticNumber = 0;
  let previousFirstSource: SourceReference | undefined;

  const noteLabels = new Map<string, number>();

  for (const note of notes) {
    /*
     * Each semicolon represents a separate source.
     *
     * Empty pieces are ignored so that accidental
     * double-semicolons don't produce empty entries.
     */
    const allSources = splitSources(
      removeLeadingNoteMarkers(note.text, note.referenceText)
    );
    const sources = isIbidSource(allSources[0] ?? "")
      ? allSources.slice(0, 1)
      : allSources;

    let reference: string;

    if (!isAutomaticReference(note.referenceText)) {
      reference = note.customReferenceText || "*";
    } else {
      automaticNumber++;
      reference = String(automaticNumber);
    }

    noteLabels.set(reference, note.index);

    const occurrences = new Map<string, number>();
    const firstSource = sources[0];
    sources.forEach((source, index) => {
      const occurrence = (occurrences.get(source) ?? 0) + 1;
      occurrences.set(source, occurrence);
      const sourceLabel =
        sources.length > 1
          ? `${reference}(${index + 1})`
          : reference;
      const directReference = readDirectReference(source);
      const isIbid = index === 0 && isIbidSource(source);
      const ibidTargetNoteIndex = isIbid && previousFirstSource
        ? previousFirstSource.isIbid &&
          previousFirstSource.ibidTargetNoteIndex
          ? previousFirstSource.ibidTargetNoteIndex
          : previousFirstSource.noteIndex
        : undefined;

      const sourceReference: SourceReference = {
        noteIndex: note.index,
        sourceIndex: index + 1,
        occurrence,
        reference: sourceLabel,
        source,
        fullNoteText: removeLeadingNoteMarkers(
          note.text,
          note.referenceText
        ),
        sourceCount: sources.length,
        directReferenceTarget: directReference?.target,
        directReferencePrefix: directReference?.prefix,
        shortNameDeclaration: readShortNameDeclaration(source),
        isIbid,
        ibidTargetNoteIndex,
        normalizedSource:
          normalizeSource(source),
      };

      references.push(sourceReference);

      if (index === 0 && firstSource === source) {
        previousFirstSource = sourceReference;
      }
    });
  }

  for (const sourceReference of references) {
    const targetLabel = sourceReference.directReferenceTarget;

    if (targetLabel) {
      const targetIndex = noteLabels.get(targetLabel);

      if (targetIndex) {
        sourceReference.directReferenceTargetIndex = targetIndex;
      } else {
        sourceReference.directReferenceTarget = `${targetLabel}:missing`;
      }
    }
  }

  return references;
}