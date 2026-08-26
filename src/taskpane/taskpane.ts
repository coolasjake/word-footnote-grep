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
  navigateToNote,
  previewFootnoteGrep,
  ItalicisedComma,
  SourceGroup,
  SourceReference,
} from "./footnotes";


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeUI();
  } else {
    setStatus("This add-in only works in Microsoft Word.", "error");
  }
});


let lastPreview: GrepMatch[] = [];
let lastItalicisedCommas: ItalicisedComma[] = [];
let lastSourceReferences: SourceReference[] = [];


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


function initializeNavigation(): void {
  $("app").addEventListener("click", (event) => {
    const target = event.target;

    if (!(target instanceof Element)) {
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


function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}


function showControlCharacters(text: string): string {
  return Array.from(text, (character) => {
    const codePoint = character.codePointAt(0) ?? 0;

    if (codePoint < 0x20 || codePoint === 0x7f) {
      return `\\u${codePoint.toString(16).padStart(4, "0")}`;
    }

    return character;
  }).join("");
}


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

function switchTab(
  tabName: "all-footnotes" | "group-sources" | "style-problems"
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


function getGroupSort(): GroupSort {
  const value = ($("group-sort") as HTMLSelectElement).value;

  return value === "alphabetical" ||
    value === "references" ||
    value === "errors"
    ? value
    : "footnote";
}

function normalizeSourceForGrouping(source: string): string {
  return source
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\s+([,.;:)])/g, "$1")
    .replace(/([(:])\s+/g, "$1")
    .toLocaleLowerCase();
}

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
  const directReferenceCount = references.filter(
    (reference) => reference.directReferenceTarget
  ).length;

  const sourceCount = new Set(
    references.map((reference) => reference.normalizedSource)
  ).size;

  summary.textContent =
    `${references.length} source reference${
      references.length === 1 ? "" : "s"
    } from ` +
    `${sourceCount} unique source${
      sourceCount === 1 ? "" : "s"
    }` +
    (mode === "grouped" && (errorCount > 0 || warningCount > 0)
      ? `; ${errorCount} error${errorCount === 1 ? "" : "s"}, ` +
        `${warningCount} warning${warningCount === 1 ? "" : "s"}`
      : "") +
    (mode === "grouped" && directReferenceCount > 0
      ? `; ${directReferenceCount} direct reference${
          directReferenceCount === 1 ? "" : "s"
        }`
      : "");

  if (mode === "grouped") {
    renderSourceGroups(groups, listId, getGroupSort());
  } else {
    renderFlatSources(references, listId);
  }
}


function buildSourceGroups(
  references: SourceReference[]
): SourceGroup[] {
  const groups: SourceGroup[] = [];
  const directGroups = new Map<string, SourceGroup>();
  const groupedReferences = new Set<SourceReference>();
  const targetedSources = new Set<string>();
  const referencesByNote = new Map<number, SourceReference[]>();

  const sourceKey = (reference: SourceReference): string =>
    `${reference.noteIndex}:${reference.sourceIndex}`;

  references.forEach((reference) => {
    const noteReferences = referencesByNote.get(reference.noteIndex) ?? [];
    noteReferences.push(reference);
    referencesByNote.set(reference.noteIndex, noteReferences);
  });

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
    if (reference.directReferenceTarget) {
      const missing = reference.directReferenceTarget.endsWith(":missing");

      if (missing) {
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

        const targetSource = matchingSources.length === 1
          ? matchingSources[0]
          : targetSources[0];
        const matchingFailed = matchingSources.length !== 1;
        const warning = matchingFailed
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

    if (
      groupedReferences.has(reference)
    ) {
      continue;
    }

    groups.push({
      source: reference.source,
      normalizedSource: reference.normalizedSource,
      references: [reference],
    });
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

function initializeUI(): void {
  initializeNavigation();

  /* Tabs */
  $("tab-all-footnotes").addEventListener("click", () => {
    switchTab("all-footnotes");
  });

  $("tab-group-sources").addEventListener("click", () => {
    switchTab("group-sources");
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