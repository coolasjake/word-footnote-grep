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

function switchTab(tabName: "search" | "commas" | "sources"): void {
  const tabs = {
    search: {
      button: $("tab-search"),
      panel: $("panel-search"),
    },
    commas: {
      button: $("tab-commas"),
      panel: $("panel-commas"),
    },
    sources: {
      button: $("tab-sources"),
      panel: $("panel-sources"),
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
        <article class="comma-card">
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

function getSelectedSourceView(): "grouped" | "flat" {
  return (
    $("source-view-grouped") as HTMLInputElement
  ).checked
    ? "grouped"
    : "flat";
}


function renderSourceGroups(groups: SourceGroup[]): void {
  const list = $("source-list");

  if (groups.length === 0) {
    list.innerHTML = `
      <div class="empty-state">
        No footnote sources found.
      </div>
    `;
    return;
  }

  list.innerHTML = groups
    .map(
      (group) => `
        <article class="source-card">
          <div class="source-header">
            <h3>${escapeHtml(group.source)}</h3>
            <span class="source-count">
              ${group.references.length}
              reference${group.references.length === 1 ? "" : "s"}
            </span>
          </div>

          <div class="source-references">
            ${group.references
              .map(
                (reference) => `
                  <span class="source-reference">
                    ${escapeHtml(reference)}
                  </span>
                `
              )
              .join("")}
          </div>
        </article>
      `
    )
    .join("");
}


function renderFlatSources(
  references: SourceReference[]
): void {
  const list = $("source-list");

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
        <article class="flat-source-card">
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
  references: SourceReference[]
): void {
  const results = $("source-results");
  const summary = $("source-summary");

  results.classList.remove("hidden");

  lastSourceReferences = references;

  const sourceCount = new Set(
    references.map((reference) => reference.normalizedSource)
  ).size;

  summary.textContent =
    `${references.length} source reference${
      references.length === 1 ? "" : "s"
    } from ` +
    `${sourceCount} unique source${
      sourceCount === 1 ? "" : "s"
    }`;

  if (getSelectedSourceView() === "grouped") {
    const groups = buildSourceGroups(references);
    renderSourceGroups(groups);
  } else {
    renderFlatSources(references);
  }
}


function buildSourceGroups(
  references: SourceReference[]
): SourceGroup[] {
  const groups = new Map<string, SourceGroup>();

  for (const reference of references) {
    const existing = groups.get(reference.normalizedSource);

    if (existing) {
      existing.references.push(reference.reference);
    } else {
      groups.set(reference.normalizedSource, {
        source: reference.source,
        normalizedSource: reference.normalizedSource,
        references: [reference.reference],
      });
    }
  }

  return Array.from(groups.values()).sort((a, b) =>
    a.source.localeCompare(b.source, undefined, {
      sensitivity: "base",
    })
  );
}


async function handleRefreshSources(): Promise<void> {
  const refreshButton =
    $("btn-refresh-sources") as HTMLButtonElement;

  refreshButton.disabled = true;
  setStatus("Reading footnotes and grouping sources…", "info");

  try {
    const references = await getFootnoteSources();

    renderSources(references);

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
  /* Tabs */
  $("tab-search").addEventListener("click", () => {
    switchTab("search");
  });

  $("tab-commas").addEventListener("click", () => {
    switchTab("commas");
  });

  $("tab-sources").addEventListener("click", () => {
    switchTab("sources");
  });


  /* Existing regex functionality */
  $("btn-preview").addEventListener("click", () => {
    void handlePreview();
  });

  $("btn-replace").addEventListener("click", () => {
    void handleReplace();
  });


  [
    "pattern",
    "replacement",
    "note-kind",
    "flag-global",
    "flag-ignore-case",
    "flag-multiline",
    "flag-dotall",
  ].forEach((id) => {
    $(id).addEventListener("input", () => {
      ($("btn-replace") as HTMLButtonElement).disabled = true;
      lastPreview = [];
    });

    $(id).addEventListener("change", () => {
      ($("btn-replace") as HTMLButtonElement).disabled = true;
      lastPreview = [];
    });
  });


  /* Italicised commas */
  $("btn-find-commas").addEventListener("click", () => {
    void handleFindItalicisedCommas();
  });

  $("btn-fix-commas").addEventListener("click", () => {
    void handleFixItalicisedCommas();
  });


  /* Sources */
  $("btn-refresh-sources").addEventListener("click", () => {
    void handleRefreshSources();
  });

  $("source-view-grouped").addEventListener("change", () => {
    if (lastSourceReferences.length > 0) {
      renderSources(lastSourceReferences);
    }
  });

  $("source-view-flat").addEventListener("change", () => {
    if (lastSourceReferences.length > 0) {
      renderSources(lastSourceReferences);
    }
  });


  void refreshNoteCounts();
}