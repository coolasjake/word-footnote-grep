import "./taskpane.css";
import {
  GrepFlags,
  GrepMatch,
  GrepOptions,
  NoteKind,
} from "./grep";

import {
  applyGrepOptions,
  findItalicisedCommas,
  fixItalicisedCommas,
  getNoteCounts,
  getFootnoteSources,
  keywordPattern,
  navigateToNote,
  previewFootnoteGrep,
  ItalicisedComma,
  SourceGroup,
  SourceReference,
} from "./footnotes";


// Wait for Office.js to finish loading before initializing the task pane.
// The UI is only initialized when the add-in is running inside Microsoft Word.
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeUI();
  } else {
    setStatus("This add-in only works in Microsoft Word.", "error");
  }
});


// Stores the most recent search preview so replacements can be applied only after a preview.
let lastPreview: GrepMatch[] = [];
// Stores the most recently detected italicised commas for the subsequent fix operation.
let lastItalicisedCommas: ItalicisedComma[] = [];
// Stores the most recently loaded source references so the grouped view can be re-sorted without rereading the document.
let lastSourceReferences: SourceReference[] = [];
let lastDetectorMatches: DetectorGroup[] = [];
const dismissedDetectorGroups = new Set<string>();


// Looks up a required DOM element by ID and throws a clear error if it is missing.
function $(id: string): HTMLElement {
  const el = document.getElementById(id);

  if (!el) {
    throw new Error(`Missing element #${id}`);
  }

  return el;
}


function setStatus(
  message: string,
  kind: "info" | "success" | "error" = "info"
): void {
  const status = $("status");

  status.textContent = message;
  status.className = `status ${kind}`;
}


// Reads navigation metadata from a result element and asks the Word API to select the corresponding note/location.
async function handleNavigation(element: HTMLElement): Promise<void> {
  const kind = element.dataset.noteKind;
  const noteIndex = Number(element.dataset.noteIndex);
  const searchText = element.dataset.searchText;
  const occurrence = Number(element.dataset.occurrence ?? "1");

  if (
    (kind !== "footnote" && kind !== "endnote") ||
    !Number.isInteger(noteIndex) ||
    noteIndex < 1 ||
    !searchText
  ) {
    return;
  }

  setStatus("Opening document location…", "info");

  try {
    await navigateToNote(
      kind,
      noteIndex,
      searchText,
      Number.isInteger(occurrence) && occurrence > 0
        ? occurrence
        : 1
    );
    setStatus("Document location selected.", "success");
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );
  }
}


// Installs delegated mouse and keyboard handlers for all dynamically rendered navigable result items.
function initializeNavigation(): void {
  $("app").addEventListener("click", (event) => {
    const target = event.target;

    if (!(target instanceof Element)) {
      return;
    }

    const dismissButton = target.closest<HTMLButtonElement>(".detector-group-dismiss");
    if (dismissButton) {
      const groupId = dismissButton.dataset.detectorGroup;
      if (groupId) {
        dismissedDetectorGroups.add(groupId);
        renderDetectorMatches();
      }
      return;
    }

    const navigable = target.closest<HTMLElement>(".navigable");

    if (navigable) {
      void handleNavigation(navigable);
    }
  });

  $("app").addEventListener("keydown", (event) => {
    if (event.key !== "Enter" && event.key !== " ") {
      return;
    }

    const target = event.target;

    if (!(target instanceof HTMLElement) ||
        !target.classList.contains("navigable")) {
      return;
    }

    event.preventDefault();
    void handleNavigation(target);
  });
}


// Reads the current search/replace controls from the task pane and converts them into GrepOptions.
function readOptions(): GrepOptions {
  const flags: GrepFlags = {
    global: ($("flag-global") as HTMLInputElement).checked,
    ignoreCase: ($("flag-ignore-case") as HTMLInputElement).checked,
    multiline: ($("flag-multiline") as HTMLInputElement).checked,
    dotAll: ($("flag-dotall") as HTMLInputElement).checked,
  };

  return {
    pattern: ($("pattern") as HTMLInputElement).value,
    replacement: ($("replacement") as HTMLInputElement).value,
    flags,
    noteKind: ($("note-kind") as HTMLSelectElement).value as NoteKind,
  };
}


// Escapes text before inserting it into generated HTML, preventing document content from being interpreted as markup.
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}


// Makes non-printing control characters visible in the UI using Unicode escape notation.
function showControlCharacters(text: string): string {
  return Array.from(text, (character) => {
    const codePoint = character.codePointAt(0) ?? 0;

    if (codePoint < 0x20 || codePoint === 0x7f) {
      return `\\u${codePoint.toString(16).padStart(4, "0")}`;
    }

    return character;
  }).join("");
}


// Formats note text for display, restoring a reference prefix when Word has returned a leading control character.
function displayNoteText(
  text: string,
  referenceText?: string
): string {
  if (referenceText && /^[\u0001-\u001f\u007f]/.test(text)) {
    return referenceText + text.slice(1);
  }

  return showControlCharacters(text);
}


/* --------------------------------------------------------------------------
   TAB HANDLING
-------------------------------------------------------------------------- */

// Activates one tab and hides the other panels, while keeping the tab ARIA state synchronized.
function switchTab(
  tabName: "all-footnotes" | "group-sources" | "group-detector" | "style-problems"
): void {
  const tabs = {
    "all-footnotes": {
      button: $("tab-all-footnotes"),
      panel: $("panel-all-footnotes"),
    },
    "group-sources": {
      button: $("tab-group-sources"),
      panel: $("panel-group-sources"),
    },
    "group-detector": {
      button: $("tab-group-detector"),
      panel: $("panel-group-detector"),
    },
    "style-problems": {
      button: $("tab-style-problems"),
      panel: $("panel-style-problems"),
    },
  };

  Object.entries(tabs).forEach(([name, tab]) => {
    const active = name === tabName;

    tab.button.classList.toggle("active", active);
    tab.button.setAttribute("aria-selected", String(active));

    tab.panel.classList.toggle("hidden", !active);
    tab.panel.toggleAttribute("hidden", !active);
  });
}


/* --------------------------------------------------------------------------
   EXISTING SEARCH / REPLACE
-------------------------------------------------------------------------- */

// Renders the search/replace preview cards and updates the match summary and Replace button state.
function renderPreview(
  matches: GrepMatch[],
  totalMatchCount: number
): void {
  const results = $("results");
  const summary = $("summary");
  const list = $("match-list");
  const replaceBtn = $("btn-replace") as HTMLButtonElement;

  if (matches.length === 0) {
    results.classList.add("hidden");
    replaceBtn.disabled = true;
    lastPreview = [];
    return;
  }

  results.classList.remove("hidden");
  replaceBtn.disabled = false;
  lastPreview = matches;

  summary.textContent =
    `${totalMatchCount} match${totalMatchCount === 1 ? "" : "es"} ` +
    `across ${matches.length} note${matches.length === 1 ? "" : "s"}`;

  list.innerHTML = matches
    .map(
      (m) => `
        <article class="match-card">
          <h3>
            ${m.noteKind === "footnote" ? "Footnote" : "Endnote"}
            ${m.noteIndex} ·
            ${m.matchCount} match${m.matchCount === 1 ? "" : "es"}
          </h3>

          <span class="diff-label">Before</span>
          <div class="diff diff-before">
            ${escapeHtml(displayNoteText(m.originalText, m.referenceText))}
          </div>

          <span class="diff-label">After</span>
          <div class="diff diff-after">
            ${escapeHtml(displayNoteText(m.newText, m.referenceText))}
          </div>
        </article>
      `
    )
    .join("");
}


// Reads the current footnote/endnote counts from Word and displays them in the task pane.
async function refreshNoteCounts(): Promise<void> {
  try {
    const counts = await getNoteCounts();

    $("note-counts").textContent =
      `${counts.footnotes} footnote${counts.footnotes === 1 ? "" : "s"}, ` +
      `${counts.endnotes} endnote${counts.endnotes === 1 ? "" : "s"} ` +
      `in document`;
  } catch (err) {
    $("note-counts").textContent = "Could not read note counts";
    console.error(err);
  }
}


// Validates the search controls, runs the preview scan, and displays its results or errors.
async function handlePreview(): Promise<void> {
  const options = readOptions();

  if (!options.pattern.trim()) {
    setStatus("Enter a search pattern.", "error");
    return;
  }

  setStatus("Scanning footnotes…", "info");
  ($("btn-preview") as HTMLButtonElement).disabled = true;

  try {
    const result = await previewFootnoteGrep(options);

    if (result.error) {
      setStatus(`Invalid regex: ${result.error}`, "error");
      renderPreview([], 0);
      return;
    }

    if (result.matches.length === 0) {
      setStatus("No matches found.", "info");
      renderPreview([], 0);
      return;
    }

    setStatus(
      `Found ${result.totalMatchCount} match${
        result.totalMatchCount === 1 ? "" : "es"
      }.`,
      "success"
    );

    renderPreview(result.matches, result.totalMatchCount);
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );

    renderPreview([], 0);
  } finally {
    ($("btn-preview") as HTMLButtonElement).disabled = false;
  }
}


// Applies the current search/replace operation using the previously generated preview as the safety check.
async function handleReplace(): Promise<void> {
  if (lastPreview.length === 0) {
    setStatus("Preview matches first.", "error");
    return;
  }

  const options = readOptions();
  const replaceBtn = $("btn-replace") as HTMLButtonElement;

  replaceBtn.disabled = true;
  setStatus("Applying replacements…", "info");

  try {
    const { applied, result } = await applyGrepOptions(options);

    if (result.error) {
      setStatus(`Error: ${result.error}`, "error");
      return;
    }

    setStatus(
      `Updated ${applied} note${applied === 1 ? "" : "s"}.`,
      "success"
    );

    renderPreview([], 0);
    await refreshNoteCounts();
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );
  } finally {
    replaceBtn.disabled = true;
  }
}


/* --------------------------------------------------------------------------
   ITALICISED COMMAS
-------------------------------------------------------------------------- */

// Renders detected italicised commas and enables/disables the Fix button based on whether any were found.
function renderItalicisedCommas(
  commas: ItalicisedComma[]
): void {
  const results = $("comma-results");
  const summary = $("comma-summary");
  const list = $("comma-list");
  const fixButton = $("btn-fix-commas") as HTMLButtonElement;

  lastItalicisedCommas = commas;

  if (commas.length === 0) {
    results.classList.remove("hidden");

    summary.textContent =
      "No italicised commas were found.";

    list.innerHTML = `
      <div class="empty-state success-state">
        No incorrectly italicised commas found.
      </div>
    `;

    fixButton.disabled = true;
    return;
  }

  results.classList.remove("hidden");
  fixButton.disabled = false;

  summary.textContent =
    `${commas.length} italicised comma${
      commas.length === 1 ? "" : "s"
    } found`;

  list.innerHTML = commas
    .map(
      (comma) => `
        <article
          class="comma-card navigable"
          data-note-kind="footnote"
          data-note-index="${comma.noteIndex}"
          data-search-text=","
          data-occurrence="${comma.occurrence}"
          tabindex="0"
          role="button"
        >
          <h3>Footnote ${comma.noteIndex}</h3>

          <div class="comma-context">
            ${escapeHtml(comma.contextBefore)}
            <span class="italic-comma">,</span>
            ${escapeHtml(comma.contextAfter)}
          </div>
        </article>
      `
    )
    .join("");
}


// Runs the italicised-comma scan and updates the UI with the results.
async function handleFindItalicisedCommas(): Promise<void> {
  const findButton =
    $("btn-find-commas") as HTMLButtonElement;

  findButton.disabled = true;
  setStatus("Scanning footnotes for italicised commas…", "info");

  try {
    const commas = await findItalicisedCommas();

    renderItalicisedCommas(commas);

    if (commas.length === 0) {
      setStatus(
        "No incorrectly italicised commas found.",
        "success"
      );
    } else {
      setStatus(
        `Found ${commas.length} italicised comma${
          commas.length === 1 ? "" : "s"
        }.`,
        "success"
      );
    }
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );

    renderItalicisedCommas([]);
  } finally {
    findButton.disabled = false;
  }
}


// Removes italic formatting from the commas found by the previous scan.
async function handleFixItalicisedCommas(): Promise<void> {
  if (lastItalicisedCommas.length === 0) {
    setStatus("Find italicised commas first.", "error");
    return;
  }

  const fixButton =
    $("btn-fix-commas") as HTMLButtonElement;

  fixButton.disabled = true;
  setStatus("Removing italic formatting from commas…", "info");

  try {
    const fixed = await fixItalicisedCommas();

    lastItalicisedCommas = [];

    renderItalicisedCommas([]);

    setStatus(
      `Fixed ${fixed} comma${fixed === 1 ? "" : "s"}.`,
      "success"
    );
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );

    fixButton.disabled = false;
  }
}


/* --------------------------------------------------------------------------
   SOURCES
-------------------------------------------------------------------------- */

type GroupSort =
  | "alphabetical"
  | "footnote"
  | "references"
  | "errors";

type DetectorQuality = "all" | "likely" | "very-likely";

interface SourceElement {
  text: string;
  originalText: string;
  words: string[];
  isKey: boolean;
  isYear: boolean;
}

interface DetectorPair {
  first: SourceReference;
  second: SourceReference;
  quality: number;
  sharesYear: boolean;
  firstMatch: string;
  secondMatch: string;
  ownerText: string;
  ownerOnFirst: boolean;
}

interface DetectorEntry {
  reference: SourceReference;
  quality: number;
  matchText?: string;
}

interface DetectorGroup {
  id: string;
  ownerText: string;
  entries: DetectorEntry[];
  score: number;
}

const detectorExcludedWordsStorageKey = "word-footnote-grep.detector-excluded-words";
const structureWords = new Set(
  "a an the and or but if then than of to in on at by for from with about as into through during before after above below between under again further once here there when where why how all any both each few more most other some such no nor not only own same so too very can will just don should now i me my mine myself you your yours yourself he him his himself she her hers herself it its itself we us our ours ourselves they them their theirs themselves this that these those is am are was were be been being do does did doing have has had having eg act etc v c s p ibid"
    .split(" ")
);

let activeExcludedWords = new Set(structureWords);

function readCustomExcludedWords(): Set<string> {
  const input = document.getElementById("detector-excluded-words") as HTMLInputElement | null;
  const enteredWords = input?.value ?? "";

  return new Set(
    enteredWords
      .split(/[\s,;]+/u)
      .map((word) => word.trim().toLocaleLowerCase())
      .filter((word) => word.length > 0)
  );
}

function loadCustomExcludedWords(): void {
  const input = document.getElementById("detector-excluded-words") as HTMLInputElement | null;
  if (!input) return;

  try {
    input.value = localStorage.getItem(detectorExcludedWordsStorageKey) ?? "";
  } catch {
    input.value = "";
  }
}

function saveCustomExcludedWords(): void {
  const input = document.getElementById("detector-excluded-words") as HTMLInputElement | null;
  if (!input) return;

  try {
    localStorage.setItem(detectorExcludedWordsStorageKey, input.value);
  } catch {
    // The detector still works when browser storage is unavailable.
  }
}

function normalizeDetectorText(text: string): string {
  keywordPattern.lastIndex = 0;
  return text
    .replace(keywordPattern, " ")
    .toLocaleLowerCase()
    .replace(/[^\p{L}\p{N}'\-\u2010\u2011\u2012\u2013\u2014]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sourceElements(source: string): SourceElement[] {
  const elements: Array<{
    text: string;
    originalText: string;
    quoted: boolean;
    isYear: boolean;
  }> = [];
  let current = "";
  let quoted = false;
  let currentWasQuoted = false;

  const addElement = (text: string, isYear = false): void => {
    const normalized = normalizeDetectorText(text);
    if (normalized) {
      elements.push({
        text: normalized,
        originalText: text.trim(),
        quoted: currentWasQuoted,
        isYear,
      });
    }
    current = "";
    currentWasQuoted = false;
  };

  for (let index = 0; index < source.length; index++) {
    const year = source.slice(index).match(/^[\[(]\s*(\d{4})\s*[\])]/);
    if (year) {
      addElement(year[1], true);
      index += year[0].length - 1;
      continue;
    }

    const character = source[index];
    const previous = source[index - 1] ?? "";
    const next = source[index + 1] ?? "";
    const isApostrophe = character === "'" || character === "\u2019";
    const isDash = /[-\u2010\u2011\u2012\u2013\u2014]/u.test(character);
    const isInternalWordPunctuation =
      (isApostrophe || isDash) &&
      /[\p{L}\p{N}]/u.test(previous) &&
      /[\p{L}\p{N}]/u.test(next);

    if (character === '"' || character === "\u2018" || character === "\u2019" ||
        character === "\u201c" || character === "\u201d" ||
        (isApostrophe && !isInternalWordPunctuation)) {
      if (!quoted && current.trim()) {
        currentWasQuoted = true;
      }
      quoted = !quoted;
      continue;
    }

    if (/[,.;:!?()[\]{}]/.test(character) ||
      (isDash && !isInternalWordPunctuation)) {
      addElement(current);
      continue;
    }

    current += character;
  }
  addElement(current);

  const firstYearIndex = elements.findIndex((element) => element.isYear);
  return elements.map((element, index) => {
    const words = element.text.match(/[\p{L}\p{N}]+/gu) ?? [];
    const isKey = !element.isYear && (
      element.quoted ||
      (index === 0 && words.length <= 5) ||
      (firstYearIndex >= 2 && index >= firstYearIndex - 2 && index < firstYearIndex)
    );

    return {
      text: element.text,
      originalText: element.originalText,
      words,
      isKey,
      isYear: element.isYear,
    };
  });
}

function usableWords(words: string[]): string[] {
  return words.filter((word) =>
    /\p{L}/u.test(word) && !activeExcludedWords.has(word)
  );
}

function isMatchableElement(element: SourceElement): boolean {
  return !element.isYear && usableWords(element.words).length > 0;
}

function partialElementMatch(
  key: SourceElement,
  other: SourceElement
): string | undefined {
  const keyWords = usableWords(key.words);
  const otherTokens = [...other.originalText.matchAll(/[\p{L}\p{N}]+/gu)].map((match) => ({
    word: match[0],
    start: match.index ?? 0,
    end: (match.index ?? 0) + match[0].length,
  }));
  const otherWords = otherTokens.map((token) => token.word);
  const minimumWords = key.words.length > 1 ? 2 : 1;

  if (keyWords.length < minimumWords) {
    return undefined;
  }

  for (let keyIndex = 0; keyIndex <= keyWords.length - minimumWords; keyIndex++) {
    const wantedWords = keyWords.slice(keyIndex, keyIndex + minimumWords);
    for (let otherIndex = 0; otherIndex < otherWords.length; otherIndex++) {
      if (otherWords[otherIndex] !== wantedWords[0]) continue;

      let nextIndex = otherIndex + 1;
      let wantedIndex = 1;
      while (wantedIndex < wantedWords.length && nextIndex < otherWords.length) {
        if (!activeExcludedWords.has(otherWords[nextIndex].toLocaleLowerCase())) {
          if (otherWords[nextIndex] !== wantedWords[wantedIndex]) break;
          wantedIndex++;
        }
        nextIndex++;
      }

      if (wantedIndex === wantedWords.length) {
        return other.originalText.slice(
          otherTokens[otherIndex].start,
          otherTokens[nextIndex - 1].end
        );
      }
    }
  }

  return undefined;
}

function containsWholeElement(key: SourceElement, other: SourceElement): boolean {
  return new RegExp(`(?:^|\\s)${key.text.replace(/[.*+?^${}()|[\\]\\]/g, "\\$&")}(?:$|\\s)`, "u")
    .test(other.text);
}

function sourcePairMatch(
  first: SourceReference,
  second: SourceReference
): DetectorPair[] {
  const firstElements = sourceElements(first.source).filter((element) => !element.isYear);
  const secondElements = sourceElements(second.source).filter((element) => !element.isYear);
  const firstYears = sourceElements(first.source).filter((element) => element.isYear).map((element) => element.text);
  const secondYears = sourceElements(second.source).filter((element) => element.isYear).map((element) => element.text);
  const sharesYear = firstYears.some((year) => secondYears.includes(year));
  const matches: DetectorPair[] = [];
  const seen = new Set<string>();
  const addMatch = (
    quality: number,
    firstMatch: string,
    secondMatch: string,
    ownerText: string,
    ownerOnFirst: boolean
  ): void => {
    const key = `${ownerText}|${firstMatch}|${secondMatch}|${quality}`;
    if (seen.has(key)) return;
    seen.add(key);
    matches.push({
      first,
      second,
      quality,
      sharesYear,
      firstMatch,
      secondMatch,
      ownerText,
      ownerOnFirst,
    });
  };

  const compareExactAndWhole = (
    keys: SourceElement[],
    others: SourceElement[],
    keyIsFirst: boolean
  ): void => {
    for (const key of keys.filter((element) => element.isKey && isMatchableElement(element))) {
      for (const other of others.filter(isMatchableElement)) {
        if (key.text === other.text) {
          const ownerOnFirst = keyIsFirst && !other.isKey;
          addMatch(
            3,
            keyIsFirst ? key.originalText : other.originalText,
            keyIsFirst ? other.originalText : key.originalText,
            key.originalText,
            ownerOnFirst
          );
        } else if (containsWholeElement(key, other)) {
          const ownerOnFirst = keyIsFirst && !other.isKey;
          addMatch(2, key.originalText, key.originalText, key.originalText, ownerOnFirst);
        }
      }
    }
  };

  compareExactAndWhole(firstElements, secondElements, true);
  compareExactAndWhole(secondElements, firstElements, false);

  const firstYear = firstYears.find((year) =>
    secondYears.includes(year) || secondElements.some((element) => element.text.includes(year))
  );
  const secondYear = secondYears.find((year) =>
    firstYears.includes(year) || firstElements.some((element) => element.text.includes(year))
  );
  const matchedYear = firstYear ?? secondYear;
  if (matchedYear) {
    addMatch(1, matchedYear, matchedYear, matchedYear, Boolean(firstYear));
  }

  const comparePartial = (
    keys: SourceElement[],
    others: SourceElement[],
    keyIsFirst: boolean
  ): void => {
    for (const key of keys.filter((element) => element.isKey && isMatchableElement(element))) {
      for (const other of others.filter(isMatchableElement)) {
        const match = partialElementMatch(key, other);
        if (match) {
          const ownerOnFirst = keyIsFirst && !other.isKey;
          addMatch(
            1,
            keyIsFirst ? key.originalText : match,
            keyIsFirst ? match : key.originalText,
            key.originalText,
            ownerOnFirst
          );
        }
      }
    }
  };

  comparePartial(firstElements, secondElements, true);
  comparePartial(secondElements, firstElements, false);

  return matches;
}

function detectorGroupKey(references: SourceReference[]): string {
  return references
    .map((reference) => `${reference.noteIndex}:${reference.sourceIndex}`)
    .sort()
    .join("|");
}

function buildDetectorMatches(
  references: SourceReference[],
  minimumQuality: number
): DetectorGroup[] {
  const linkedGroups = buildSourceGroups(references);
  const groupByReference = new Map<SourceReference, SourceGroup>();
  linkedGroups.forEach((group) => group.references.forEach((reference) => groupByReference.set(reference, group)));
  const groupsByOwner = new Map<string, DetectorGroup>();

  const addEntry = (
    group: DetectorGroup,
    reference: SourceReference,
    quality: number,
    matchText: string
  ): void => {
    const existing = group.entries.find((entry) => entry.reference === reference);
    if (existing) {
      if (existing.quality < quality) {
        existing.quality = quality;
        existing.matchText = matchText;
      }
      return;
    }

    group.entries.push({
      reference,
      quality,
      matchText,
    });
  };

  for (let firstIndex = 0; firstIndex < references.length; firstIndex++) {
    for (let secondIndex = firstIndex + 1; secondIndex < references.length; secondIndex++) {
      const first = references[firstIndex];
      const second = references[secondIndex];
      const firstGroup = groupByReference.get(first);
      const secondGroup = groupByReference.get(second);
      if (firstGroup && firstGroup === secondGroup) continue;

      for (const pair of sourcePairMatch(first, second)) {
        if (pair.quality < minimumQuality) continue;
        const ownerKey = normalizeDetectorText(pair.ownerText);
        let group = groupsByOwner.get(ownerKey);
        if (!group) {
          group = {
            id: `${ownerKey}|${first.noteIndex}:${first.sourceIndex}|${second.noteIndex}:${second.sourceIndex}`,
            ownerText: pair.ownerText,
            entries: [],
            score: 0,
          };
          groupsByOwner.set(ownerKey, group);
        }

        const ownerReference = pair.ownerOnFirst ? first : second;
        const ownerMatch = pair.ownerText;
        addEntry(
          group,
          first,
          pair.quality,
          ownerReference === first ? ownerMatch : pair.firstMatch
        );
        addEntry(
          group,
          second,
          pair.quality,
          ownerReference === second ? ownerMatch : pair.secondMatch
        );
        group.score = Math.max(
          group.score,
          pair.quality * 10 + Number(pair.sharesYear)
        );
      }
    }
  }

  return [...groupsByOwner.values()]
    .filter((group) => group.entries.length > 1)
    .map((group) => {
      group.id = `${group.ownerText.toLocaleLowerCase()}|${detectorGroupKey(group.entries.map((entry) => entry.reference))}`;
      return group;
    })
    .sort((first, second) => second.score - first.score);
}

function selectedDetectorQuality(): DetectorQuality {
  const selected = document.querySelector<HTMLInputElement>(
    'input[name="detector-quality"]:checked'
  );

  return (selected?.value as DetectorQuality) ?? "all";
}

function detectorGroupIsVisible(group: DetectorGroup): boolean {
  const selected = selectedDetectorQuality();
  const minimumQuality = selected === "very-likely" ? 3 : selected === "likely" ? 2 : 1;
  return group.score >= minimumQuality * 10;
}

function visibleDetectorEntries(group: DetectorGroup): DetectorEntry[] {
  const selected = selectedDetectorQuality();
  const minimumQuality = selected === "very-likely" ? 3 : selected === "likely" ? 2 : 1;
  return group.entries.filter((entry) =>
    entry.quality >= minimumQuality
  );
}

function detectorQualityLabel(quality: number): string {
  return quality >= 3 ? "Very Likely" : quality === 2 ? "Likely" : "Possible";
}

function renderMatchedSource(source: string, matchText?: string): string {
  if (!matchText) return escapeHtml(source);
  const directIndex = source.toLocaleLowerCase().indexOf(matchText.toLocaleLowerCase());
  let matchStart = directIndex;
  let matchEnd = directIndex < 0 ? -1 : directIndex + matchText.length;

  if (matchStart < 0) {
    const sourceTokens = [...source.matchAll(/[\p{L}\p{N}]+/gu)].map((match) => ({
      word: match[0].toLocaleLowerCase(),
      start: match.index ?? 0,
      end: (match.index ?? 0) + match[0].length,
    }));
    const wantedWords = normalizeDetectorText(matchText)
      .match(/[\p{L}\p{N}]+/gu) ?? [];

    for (let start = 0; start < sourceTokens.length; start++) {
      let wantedIndex = 0;
      let tokenIndex = start;
      while (tokenIndex < sourceTokens.length && wantedIndex < wantedWords.length) {
        const token = sourceTokens[tokenIndex];
        if (!activeExcludedWords.has(token.word)) {
          if (token.word !== wantedWords[wantedIndex]) break;
          wantedIndex++;
        }
        tokenIndex++;
      }

      if (wantedIndex === wantedWords.length && wantedWords.length > 0) {
        matchStart = sourceTokens[start].start;
        matchEnd = sourceTokens[tokenIndex - 1].end;
        break;
      }
    }
  }

  if (matchStart < 0 || matchEnd < 0) return escapeHtml(source);
  return escapeHtml(source.slice(0, matchStart)) +
    `<strong>**${escapeHtml(source.slice(matchStart, matchEnd))}**</strong>` +
    escapeHtml(source.slice(matchEnd));
}

function renderDetectorMatches(): void {
  const results = $("group-detector-results");
  const list = $("group-detector-list");
  const visibleGroups = lastDetectorMatches.filter((group) =>
    !dismissedDetectorGroups.has(group.id) && detectorGroupIsVisible(group)
  );
  results.classList.remove("hidden");
  $("group-detector-summary").textContent =
    `${visibleGroups.length} possible match group${visibleGroups.length === 1 ? "" : "s"}`;

  if (visibleGroups.length === 0) {
    list.innerHTML = '<div class="empty-state">No possible source matches found.</div>';
    return;
  }

  list.innerHTML = visibleGroups.map((group) => `
    <article class="detector-group">
      <button class="detector-group-dismiss" type="button" data-detector-group="${group.id}" aria-label="Dismiss match group">×</button>
      <h3 class="detector-owner">${renderMatchedSource(group.ownerText)}</h3>
      ${visibleDetectorEntries(group).map((entry) => `
        <div class="detector-match">
          <div
            class="detector-source navigable"
            data-note-kind="footnote"
            data-note-index="${entry.reference.noteIndex}"
            data-search-text="${escapeHtml(entry.reference.source)}"
            data-occurrence="${entry.reference.occurrence}"
            tabindex="0"
            role="button"
          >
            <span class="detector-reference">${escapeHtml(entry.reference.reference)}</span>
            ${renderMatchedSource(entry.reference.source, entry.matchText)}
          </div>
          <span class="detector-quality">${detectorQualityLabel(entry.quality)}</span>
        </div>
      `).join("")}
    </article>
  `).join("");
}

async function handleRefreshDetector(): Promise<void> {
  const button = $("btn-refresh-group-detector") as HTMLButtonElement;
  button.disabled = true;
  setStatus("Reading footnotes and finding similar sources…", "info");
  dismissedDetectorGroups.clear();

  try {
    saveCustomExcludedWords();
    activeExcludedWords = new Set([
      ...structureWords,
      ...readCustomExcludedWords(),
    ]);
    const selected = selectedDetectorQuality();
    const minimumQuality = selected === "very-likely" ? 3 : selected === "likely" ? 2 : 1;
    lastDetectorMatches = buildDetectorMatches(await getFootnoteSources(), minimumQuality);
    renderDetectorMatches();
    setStatus(`Found ${lastDetectorMatches.length} possible match group${lastDetectorMatches.length === 1 ? "" : "s"}.`, "success");
  } catch (err) {
    setStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    button.disabled = false;
  }
}


// Reads and validates the selected source-group sort order, falling back to footnote order.
function getGroupSort(): GroupSort {
  const value = ($("group-sort") as HTMLSelectElement).value;

  return value === "alphabetical" ||
    value === "references" ||
    value === "errors"
    ? value
    : "footnote";
}

// Normalizes source text for grouping comparisons by trimming whitespace and punctuation spacing and ignoring case.
function normalizeSourceForGrouping(source: string): string {
  return source
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(:])\s+/g, "$1")
    .toLocaleLowerCase();
}

// Normalizes a declared short source name so it can be compared consistently.
function normalizeShortName(name: string): string {
  return name.trim().replace(/\s+/g, " ").toLocaleLowerCase();
}

// Finds an earlier short-name declaration that appears to match the supplied source text.
function findShortNameDeclaration(
  source: string,
  declarations: Map<string, SourceReference>
): SourceReference | undefined {
  const normalizedSource = normalizeShortName(source);

  for (const [name, declaration] of declarations) {
    if (
      normalizedSource === name ||
      normalizedSource.startsWith(`${name} `) ||
      normalizedSource.startsWith(`${name},`) ||
      normalizedSource.startsWith(`${name}.`) ||
      normalizedSource.startsWith(`${name}:`)
    ) {
      return declaration;
    }
  }

  return undefined;
}

// Renders grouped sources, applying the requested sort order and showing any grouping warnings/errors.
function renderSourceGroups(
  groups: SourceGroup[],
  listId: string,
  sort: GroupSort
): void {
  const list = $(listId);

  if (groups.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        No footnote sources found.
      </div>
    `;
    return;
  }

  const orderedGroups = [...groups].sort((a, b) => {
    if (sort === "errors") {
      const problemOrder =
        (Number(Boolean(b.error)) * 2 + Number(Boolean(b.warning))) -
        (Number(Boolean(a.error)) * 2 + Number(Boolean(a.warning)));

      if (problemOrder !== 0) {
        return problemOrder;
      }
    }

    if (sort === "alphabetical") {
      const alphabeticalOrder = a.source.trimStart().localeCompare(
        b.source.trimStart(),
        undefined,
        { sensitivity: "base" }
      );

      if (alphabeticalOrder !== 0) {
        return alphabeticalOrder;
      }
    }

    if (sort === "references") {
      const referenceOrder =
        b.references.length - a.references.length;

      if (referenceOrder !== 0) {
        return referenceOrder;
      }
    }

    return a.references[0].noteIndex - b.references[0].noteIndex;
  });

  list.innerHTML = orderedGroups
    .map(
      (group) => `
        <article class="source-card">
          <div class="source-header">
            <h3
              class="navigable"
              data-note-kind="footnote"
              data-note-index="${group.references[0].noteIndex}"
              data-search-text="${escapeHtml(group.source)}"
              data-occurrence="${group.references[0].occurrence}"
              tabindex="0"
              role="button"
            >${escapeHtml(group.source)}</h3>
            <span class="source-count">
              ${group.references.length}
              reference${group.references.length === 1 ? "" : "s"}
            </span>
          </div>

          <div class="source-references">
            ${group.references
              .map(
                (reference) => `
                  <span
                    class="source-reference navigable"
                    data-note-kind="footnote"
                    data-note-index="${reference.noteIndex}"
                    data-search-text="${escapeHtml(reference.source)}"
                    data-occurrence="${reference.occurrence}"
                    tabindex="0"
                    role="button"
                  >
                    ${escapeHtml(reference.reference)}
                  </span>
                `
              )
              .join("")}
          </div>
          ${group.error
            ? `<p class="group-problem group-error">${escapeHtml(group.error)}</p>`
            : group.warning
              ? `<p class="group-problem group-warning">${escapeHtml(group.warning)}</p>`
              : ""}
        </article>
      `
    )
    .join("");
}


// Renders the ungrouped list of source references in document/footnote order.
function renderFlatSources(
  references: SourceReference[],
  listId: string
): void {
  const list = $(listId);

  if (references.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        No footnote sources found.
      </div>
    `;
    return;
  }

  list.innerHTML = references
    .map(
      (source) => `
        <article
          class="flat-source-card navigable"
          data-note-kind="footnote"
          data-note-index="${source.noteIndex}"
          data-search-text="${escapeHtml(source.source)}"
          data-occurrence="${source.occurrence}"
          tabindex="0"
          role="button"
        >
          <span class="flat-reference">
            ${escapeHtml(source.reference)}
          </span>

          <span class="flat-source">
            ${escapeHtml(source.source)}
          </span>
        </article>
      `
    )
    .join("");
}


// Chooses the grouped or flat renderer, calculates summary statistics, and stores the loaded references.
function renderSources(
  references: SourceReference[],
  mode: "grouped" | "flat",
  resultsId: string,
  summaryId: string,
  listId: string
): void {
  const results = $(resultsId);
  const summary = $(summaryId);

  results.classList.remove("hidden");

  lastSourceReferences = references;

  const groups = mode === "grouped"
    ? buildSourceGroups(references)
    : [];
  const errorCount = groups.filter((group) => group.error).length;
  const warningCount = groups.filter((group) => group.warning).length;
  const footnoteCount = new Set(
    references.map((reference) => reference.noteIndex)
  ).size;

  summary.textContent = mode === "grouped"
    ? `${groups.length} source${groups.length === 1 ? "" : "s"}, ` +
      `${footnoteCount} footnote${footnoteCount === 1 ? "" : "s"}, ` +
      `${errorCount} error${errorCount === 1 ? "" : "s"}, ` +
      `${warningCount} warning${warningCount === 1 ? "" : "s"}`
    : `${references.length} source${references.length === 1 ? "" : "s"}, ` +
      `${footnoteCount} footnote${footnoteCount === 1 ? "" : "s"}`;

  if (mode === "grouped") {
    renderSourceGroups(groups, listId, getGroupSort());
  } else {
    renderFlatSources(references, listId);
  }
}


// Builds logical source groups by resolving direct references, Ibid references, short names, and unresolved references.
// The function also records warnings/errors so the UI can flag ambiguous or broken relationships.
function buildSourceGroups(
  references: SourceReference[]
): SourceGroup[] {
  // Final collection of source groups that will eventually be displayed.
  const groups: SourceGroup[] = [];
  // Maps a target source to the group being built around that source.
  const directGroups = new Map<string, SourceGroup>();
  // Prevents a reference from being assigned to more than one group.
  const groupedReferences = new Set<SourceReference>();
  // Tracks source declarations that have been targeted by another reference.
  const targetedSources = new Set<string>();
  // Indexes source references by footnote number so direct/Ibid references can find their targets quickly.
  const referencesByNote = new Map<number, SourceReference[]>();
  // Stores declared short names and the source reference where each was declared.
  const shortNameDeclarations = new Map<string, SourceReference>();

  // Produces a stable key for a particular source occurrence within a footnote.
  const sourceKey = (reference: SourceReference): string =>
    `${reference.noteIndex}:${reference.sourceIndex}`;

  references.forEach((reference) => {
    const noteReferences = referencesByNote.get(reference.noteIndex) ?? [];
    noteReferences.push(reference);
    referencesByNote.set(reference.noteIndex, noteReferences);
  });

  // Adds a reference to a target source group, creating that group when necessary.
  const addDirectMember = (
    target: SourceReference,
    member: SourceReference,
    warning?: string,
  ): void => {
    const groupKey = `${target.noteIndex}:${target.sourceIndex}`;
    let group = directGroups.get(groupKey);

    if (!group) {
      group = {
        source: target.source,
        normalizedSource: target.normalizedSource,
        references: [],
      };
      directGroups.set(groupKey, group);
      groups.push(group);
    }

    if (!group.references.includes(member)) {
      group.references.push(member);
    }
    if (warning) {
      group.warning = warning;
    }
    groupedReferences.add(member);
  };

  for (const reference of references) {
    const declarationName = reference.shortNameDeclaration;

    if (groupedReferences.has(reference)) {
      if (declarationName) {
        shortNameDeclarations.set(
          normalizeShortName(declarationName),
          reference
        );
      }
      continue;
    }

    if (reference.isIbid) {
      const target = reference.ibidTargetNoteIndex
        ? referencesByNote.get(reference.ibidTargetNoteIndex)?.[0]
        : undefined;

      if (
        target &&
        !target.directReferenceTarget?.endsWith(":missing")
      ) {
        const warning = target.sourceCount > 1
          ? "Warning: Ibid refers to a footnote containing multiple sources."
          : undefined;
        targetedSources.add(sourceKey(target));
        addDirectMember(target, target, warning);
        addDirectMember(target, reference, warning);
      } else {
        groups.push({
          source: reference.source,
          normalizedSource: reference.normalizedSource,
          references: [reference],
          warning: "Warning: Ibid follows an unresolved footnote.",
        });
        groupedReferences.add(reference);
      }
      continue;
    }

    if (reference.directReferenceTarget) {
      const missing = reference.directReferenceTarget.endsWith(":missing");

      if (missing) {
        const shortNameDeclaration = reference.directReferencePrefix
          ? shortNameDeclarations.get(
              normalizeShortName(reference.directReferencePrefix)
            )
          : undefined;

        if (shortNameDeclaration) {
          targetedSources.add(sourceKey(shortNameDeclaration));
          addDirectMember(
            shortNameDeclaration,
            shortNameDeclaration,
            `Mismatched short name found in footnote ${reference.directReferenceTarget.slice(0, -8)}.`
          );
          addDirectMember(
            shortNameDeclaration,
            reference,
            `Mismatched short name found in footnote ${reference.directReferenceTarget.slice(0, -8)}.`
          );
          continue;
        }

        groups.push({
          source: reference.fullNoteText,
          normalizedSource: reference.normalizedSource,
          references: [reference],
          error: `Error: direct reference (n ${reference.directReferenceTarget.slice(0, -8)}) does not match a footnote number.`,
        });
        groupedReferences.add(reference);
        continue;
      }

      const target = reference.directReferenceTargetIndex
        ? referencesByNote.get(reference.directReferenceTargetIndex)?.[0]
        : undefined;

      if (target) {
        const targetSources = referencesByNote.get(target.noteIndex) ?? [];
        const prefix = normalizeSourceForGrouping(
          reference.directReferencePrefix ?? ""
        );
        const matchingSources = targetSources.filter((candidate) =>
          normalizeSourceForGrouping(candidate.source).includes(prefix)
        );

        const matchingFailed = matchingSources.length !== 1;
        const shortNameDeclaration = matchingFailed &&
          reference.directReferencePrefix
          ? shortNameDeclarations.get(
              normalizeShortName(reference.directReferencePrefix)
            )
          : undefined;
        const targetSource = shortNameDeclaration ?? (
          matchingSources.length === 1
            ? matchingSources[0]
            : targetSources[0]
        );
        const warning = shortNameDeclaration
          ? `Mismatched short name found in footnote ${reference.directReferenceTarget}.`
          : matchingFailed
            ? "Automatic matching failed."
            : undefined;

        if (targetSource) {
          targetedSources.add(sourceKey(targetSource));
          addDirectMember(targetSource, targetSource, warning);
          addDirectMember(targetSource, reference, warning);
        }
        continue;
      }
    }

    const shortNameDeclaration = declarationName
      ? undefined
      : findShortNameDeclaration(reference.source, shortNameDeclarations);

    if (shortNameDeclaration) {
      targetedSources.add(sourceKey(shortNameDeclaration));
      const warning =
        `Mismatched short name found in footnote ${reference.noteIndex}.`;
      addDirectMember(shortNameDeclaration, shortNameDeclaration, warning);
      addDirectMember(shortNameDeclaration, reference, warning);
      continue;
    }

    groups.push({
      source: reference.source,
      normalizedSource: reference.normalizedSource,
      references: [reference],
    });

    if (declarationName) {
      shortNameDeclarations.set(
        normalizeShortName(declarationName),
        reference
      );
    }
  }

  return groups
    .filter(
      (group) =>
        group.error ||
        group.references.length !== 1 ||
        !targetedSources.has(sourceKey(group.references[0]))
    )
    .sort(
    (a, b) =>
      a.references[0].noteIndex - b.references[0].noteIndex
    );
}


// Reads source references from Word and refreshes either the flat or grouped source view.
async function handleRefreshSources(
  mode: "grouped" | "flat"
): Promise<void> {
  const buttonId = mode === "grouped"
    ? "btn-refresh-group-sources"
    : "btn-refresh-all-footnotes";
  const refreshButton = $(buttonId) as HTMLButtonElement;

  refreshButton.disabled = true;
  setStatus("Reading footnotes and grouping sources…", "info");

  try {
    const references = await getFootnoteSources();

    const panelIds = mode === "grouped"
      ? {
          results: "group-sources-results",
          summary: "group-sources-summary",
          list: "group-sources-list",
        }
      : {
          results: "all-footnotes-results",
          summary: "all-footnotes-summary",
          list: "all-footnotes-list",
        };

    renderSources(
      references,
      mode,
      panelIds.results,
      panelIds.summary,
      panelIds.list
    );

    const uniqueSources = new Set(
      references.map((reference) => reference.normalizedSource)
    ).size;

    setStatus(
      `Found ${references.length} source reference${
        references.length === 1 ? "" : "s"
      } across ${uniqueSources} source${
        uniqueSources === 1 ? "" : "s"
      }.`,
      "success"
    );
  } catch (err) {
    setStatus(
      err instanceof Error ? err.message : String(err),
      "error"
    );
  } finally {
    refreshButton.disabled = false;
  }
}


/* --------------------------------------------------------------------------
   INITIALISATION
-------------------------------------------------------------------------- */

// Wires all task-pane controls to their handlers and performs the initial document count refresh.
function initializeUI(): void {
  initializeNavigation();
  loadCustomExcludedWords();

  /* Tabs */
  $("tab-all-footnotes").addEventListener("click", () => {
    switchTab("all-footnotes");
  });

  $("tab-group-sources").addEventListener("click", () => {
    switchTab("group-sources");
  });

  $("tab-group-detector").addEventListener("click", () => {
    switchTab("group-detector");
  });

  $("tab-style-problems").addEventListener("click", () => {
    switchTab("style-problems");
  });


  /* Italicised commas */
  $("btn-find-commas").addEventListener("click", () => {
    void handleFindItalicisedCommas();
  });

  $("btn-fix-commas").addEventListener("click", () => {
    void handleFixItalicisedCommas();
  });


  /* Source views */
  $("btn-refresh-all-footnotes").addEventListener("click", () => {
    void handleRefreshSources("flat");
  });

  $("btn-refresh-group-sources").addEventListener("click", () => {
    void handleRefreshSources("grouped");
  });

  $("btn-refresh-group-detector").addEventListener("click", () => {
    void handleRefreshDetector();
  });

  ["detector-show-all", "detector-show-likely", "detector-show-very-likely"]
    .forEach((id) => {
      $(id).addEventListener("change", () => {
        if (lastDetectorMatches.length > 0) {
          renderDetectorMatches();
        }
      });
    });

  $("group-sort").addEventListener("change", () => {
    if (lastSourceReferences.length > 0) {
      renderSources(
        lastSourceReferences,
        "grouped",
        "group-sources-results",
        "group-sources-summary",
        "group-sources-list"
      );
    }
  });


  void refreshNoteCounts();
}