import { getCurrentEmailContext, EmailContext } from "../services/office-item.service";
import { getMailFolders, moveItemToFolder, MailFolder } from "../services/ews.service";
import { suggestFolders, FolderSuggestion } from "../services/folder-suggester.service";

// ---------- Module state ----------
let currentContext: EmailContext | null = null;
let allFolders: MailFolder[] = [];
let selectedFolder: MailFolder | null = null;

// ---------- Element accessors (typed, fail loudly if HTML/TS drift apart) ----------
function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) {
    throw new Error(`Expected element #${id} to exist in taskpane.html`);
  }
  return node as T;
}

// =====================================================================
// Office.js lifecycle entry point.
// Per Microsoft guidance, ALL Office.js API calls must wait for
// Office.onReady (or Office.initialize) to fire before executing.
// =====================================================================
Office.onReady((info: { host: Office.HostType; platform: Office.PlatformType }) => {
  try {
    el("app-loading").setAttribute("hidden", "true");

    if (info.host !== Office.HostType.Outlook) {
      // This add-in's manifest only targets Mailbox, so this should be
      // unreachable in practice, but we guard defensively anyway.
      showFatalError("This add-in only runs inside Outlook.");
      return;
    }

    el("app-body").style.display = "flex";
    wireStaticEventHandlers();

    initialize().catch((err: unknown) => {
      showFatalError(errorMessage(err));
    });
  } catch (err) {
    // Office.onReady's callback itself must never throw uncaught —
    // fall back to a raw alert since the DOM may not be usable yet.
    // eslint-disable-next-line no-console
    console.error("Fatal error during Office.onReady:", err);
  }
});

// =====================================================================
// Main flow
// =====================================================================
async function initialize(): Promise<void> {
  resetUi();
  setLoading(true, "Reading the open email…");

  try {
    currentContext = await getCurrentEmailContext();
    renderEmailSummary(currentContext);

    setLoading(true, "Looking up your mail folders…");
    allFolders = await getMailFolders();

    if (allFolders.length === 0) {
      showEmptyState("No folders were found in this mailbox.");
      return;
    }

    populateManualPicker(allFolders);
    el("manual-picker-section").removeAttribute("hidden");

    setLoading(true, "Matching folders to this email…");
    const suggestions = suggestFolders(currentContext, allFolders);

    if (suggestions.length === 0) {
      showEmptyState("We couldn't find a confident match — choose a folder manually below.");
      return;
    }

    renderSuggestions(suggestions);
  } catch (err) {
    showFatalError(errorMessage(err));
  } finally {
    setLoading(false);
  }
}

async function fileEmail(): Promise<void> {
  if (!currentContext || !selectedFolder) {
    return;
  }

  const button = el<HTMLButtonElement>("file-button");
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = "Filing…";
  hideBanner("status-error");
  hideBanner("status-success");

  try {
    await moveItemToFolder(currentContext.itemId, selectedFolder.id);
    showSuccess(`Moved to "${selectedFolder.displayName}".`);
    button.textContent = "Filed ✓";
  } catch (err) {
    showError(errorMessage(err));
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

// =====================================================================
// Rendering
// =====================================================================
function renderEmailSummary(context: EmailContext): void {
  el("email-sender").textContent = context.senderName
    ? `${context.senderName} <${context.senderEmail}>`
    : context.senderEmail;
  el("email-subject").textContent = context.subject;
  el("email-summary").removeAttribute("hidden");
}

function renderSuggestions(suggestions: FolderSuggestion[]): void {
  const list = el("suggestions-list");
  list.innerHTML = "";

  const maxScore = Math.max(...suggestions.map((s) => s.score), 1);

  suggestions.forEach((suggestion, index) => {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "suggestion-card";
    card.setAttribute("role", "radio");
    card.setAttribute("aria-checked", "false");

    const confidencePct = Math.round((suggestion.score / maxScore) * 100);

    const reasonsHtml = suggestion.reasons
      .slice(0, 3)
      .map((r) => `<li>${escapeHtml(r)}</li>`)
      .join("");

    card.innerHTML = `
      <span class="suggestion-card__radio" aria-hidden="true"></span>
      <span class="suggestion-card__body">
        <span class="suggestion-card__top-row">
          <span class="suggestion-card__name">${escapeHtml(suggestion.folder.displayName)}</span>
          <span class="suggestion-card__score">${confidencePct}% match</span>
        </span>
        <ul class="suggestion-card__reasons">${reasonsHtml}</ul>
      </span>
    `;

    card.addEventListener("click", () => selectFolder(suggestion.folder, card));

    list.appendChild(card);

    // Pre-select the top suggestion so the user can file in one click.
    if (index === 0) {
      selectFolder(suggestion.folder, card);
    }
  });

  el("suggestions-section").removeAttribute("hidden");
}

function populateManualPicker(folders: MailFolder[]): void {
  const select = el<HTMLSelectElement>("manual-folder-select");
  // Keep the placeholder option, drop any previously-added folder options.
  while (select.options.length > 1) {
    select.remove(1);
  }

  folders
    .slice()
    .sort((a, b) => a.displayName.localeCompare(b.displayName))
    .forEach((folder) => {
      const option = document.createElement("option");
      option.value = folder.id;
      option.textContent = folder.displayName;
      select.appendChild(option);
    });
}

function selectFolder(folder: MailFolder, cardEl?: HTMLElement): void {
  selectedFolder = folder;

  // Clear selection state on all suggestion cards, then mark the chosen one.
  const cards = Array.prototype.slice.call(document.querySelectorAll(".suggestion-card")) as HTMLElement[];
  cards.forEach((c) => c.setAttribute("aria-checked", "false"));
  if (cardEl) {
    cardEl.setAttribute("aria-checked", "true");
  }

  // Keep the manual dropdown in sync so the two selection UIs never disagree.
  const select = el<HTMLSelectElement>("manual-folder-select");
  select.value = folder.id;

  const button = el<HTMLButtonElement>("file-button");
  button.disabled = false;
  button.textContent = `File to "${folder.displayName}"`;
  hideBanner("status-success");
}

// =====================================================================
// UI event wiring (attached once, outside the async initialize() flow)
// =====================================================================
function wireStaticEventHandlers(): void {
  el<HTMLButtonElement>("file-button").addEventListener("click", () => {
    fileEmail().catch((err) => showError(errorMessage(err)));
  });

  el<HTMLButtonElement>("manual-picker-toggle").addEventListener("click", () => {
    const body = el("manual-picker-body");
    const toggle = el<HTMLButtonElement>("manual-picker-toggle");
    const isHidden = body.hasAttribute("hidden");
    if (isHidden) {
      body.removeAttribute("hidden");
      toggle.setAttribute("aria-expanded", "true");
    } else {
      body.setAttribute("hidden", "true");
      toggle.setAttribute("aria-expanded", "false");
    }
  });

  el<HTMLSelectElement>("manual-folder-select").addEventListener("change", (event) => {
    const select = event.target as HTMLSelectElement;
    const folder = allFolders.find((f) => f.id === select.value);
    if (folder) {
      selectFolder(folder);
    }
  });
}

// =====================================================================
// Status / loading helpers
// =====================================================================
function resetUi(): void {
  el("email-summary").setAttribute("hidden", "true");
  el("suggestions-section").setAttribute("hidden", "true");
  el("suggestions-list").innerHTML = "";
  el("empty-state").setAttribute("hidden", "true");
  el("manual-picker-section").setAttribute("hidden", "true");
  el("manual-picker-body").setAttribute("hidden", "true");
  hideBanner("status-error");
  hideBanner("status-success");
  const button = el<HTMLButtonElement>("file-button");
  button.disabled = true;
  button.textContent = "File email";
}

function setLoading(isLoading: boolean, text?: string): void {
  const banner = el("status-loading");
  if (isLoading) {
    if (text) {
      el("status-loading-text").textContent = text;
    }
    banner.removeAttribute("hidden");
  } else {
    banner.setAttribute("hidden", "true");
  }
}

function showEmptyState(message: string): void {
  el("empty-state-text").textContent = message;
  el("empty-state").removeAttribute("hidden");
}

function showError(message: string): void {
  const banner = el("status-error");
  banner.textContent = message;
  banner.removeAttribute("hidden");
}

function showSuccess(message: string): void {
  const banner = el("status-success");
  banner.textContent = message;
  banner.removeAttribute("hidden");
}

function hideBanner(id: string): void {
  el(id).setAttribute("hidden", "true");
}

function showFatalError(message: string): void {
  setLoading(false);
  showError(message);
}

function errorMessage(err: unknown): string {
  if (err instanceof Error) {
    return err.message;
  }
  return "An unexpected error occurred.";
}

function escapeHtml(value: string): string {
  const div = document.createElement("div");
  div.textContent = value;
  return div.innerHTML;
}
