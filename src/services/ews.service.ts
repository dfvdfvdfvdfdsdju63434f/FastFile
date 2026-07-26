/**
 * All mailbox folder listing / moving goes through Office.js's built-in EWS
 * bridge (Office.context.mailbox.makeEwsRequestAsync). This requires the
 * ReadWriteMailbox permission in the manifest but needs NO external backend,
 * NO Azure AD app registration, and NO OAuth flow — Outlook signs and scopes
 * the request to the current user's own mailbox automatically.
 *
 * Known limitation: makeEwsRequestAsync is not available on every Outlook
 * client (some "new Outlook" builds and mobile clients are Graph-only).
 * getMailFolders() / moveItemToFolder() below throw a clear, catchable error
 * in that case rather than failing silently — see README "Limitations".
 */

export interface MailFolder {
  id: string;
  displayName: string;
  totalCount?: number;
}

const NS = {
  soap: "http://schemas.xmlsoap.org/soap/envelope/",
  t: "http://schemas.microsoft.com/exchange/services/2006/types",
  m: "http://schemas.microsoft.com/exchange/services/2006/messages",
};

// Folders that are rarely useful "file this email" destinations.
const EXCLUDED_FOLDER_NAMES = new Set([
  "Sync Issues",
  "Conversation History",
  "Junk Email",
  "Deleted Items",
  "Outbox",
  "Drafts",
  "RSS Feeds",
]);

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function buildFindFoldersRequest(): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:soap="${NS.soap}" xmlns:t="${NS.t}" xmlns:m="${NS.m}">` +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\" /></soap:Header>" +
    "<soap:Body>" +
    '<m:FindFolder Traversal="Deep">' +
    "<m:FolderShape>" +
    "<t:BaseShape>Default</t:BaseShape>" +
    "<t:AdditionalProperties>" +
    '<t:FieldURI FieldURI="folder:TotalCount" />' +
    "</t:AdditionalProperties>" +
    "</m:FolderShape>" +
    "<m:ParentFolderIds>" +
    '<t:DistinguishedFolderId Id="msgfolderroot" />' +
    "</m:ParentFolderIds>" +
    "</m:FindFolder>" +
    "</soap:Body>" +
    "</soap:Envelope>"
  );
}

function buildMoveItemRequest(itemId: string, folderId: string): string {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    `<soap:Envelope xmlns:soap="${NS.soap}" xmlns:t="${NS.t}" xmlns:m="${NS.m}">` +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\" /></soap:Header>" +
    "<soap:Body>" +
    "<m:MoveItem>" +
    "<m:ToFolderId>" +
    `<t:FolderId Id="${escapeXml(folderId)}" />` +
    "</m:ToFolderId>" +
    "<m:ItemIds>" +
    `<t:ItemId Id="${escapeXml(itemId)}" />` +
    "</m:ItemIds>" +
    "</m:MoveItem>" +
    "</soap:Body>" +
    "</soap:Envelope>"
  );
}

function makeEwsRequest(requestXml: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const mailbox = Office.context.mailbox as any;
    if (!mailbox || typeof mailbox.makeEwsRequestAsync !== "function") {
      reject(
        new Error(
          "This Outlook client doesn't support direct folder access (EWS). " +
            "Try Outlook on Windows/Mac (classic) or Outlook on the web."
        )
      );
      return;
    }
    mailbox.makeEwsRequestAsync(requestXml, (result: Office.AsyncResult<string>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        const message = (result.error && result.error.message) || "The mailbox request failed.";
        reject(new Error(message));
      }
    });
  });
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml");
}

function assertNoEwsError(doc: Document): void {
  const responseCodeNode = doc.getElementsByTagNameNS(NS.m, "ResponseCode")[0];
  const responseCode = responseCodeNode && responseCodeNode.textContent;
  if (responseCode && responseCode !== "NoError") {
    const messageNode = doc.getElementsByTagNameNS(NS.m, "MessageText")[0];
    const message = (messageNode && messageNode.textContent) || `Mailbox server returned: ${responseCode}`;
    throw new Error(message);
  }
}

/** Fetches all mail folders in the mailbox (excluding system/noise folders). */
export async function getMailFolders(): Promise<MailFolder[]> {
  const xml = await makeEwsRequest(buildFindFoldersRequest());
  const doc = parseXml(xml);
  assertNoEwsError(doc);

  const folderNodes = Array.prototype.slice.call(doc.getElementsByTagNameNS(NS.t, "Folder"));

  const folders: MailFolder[] = [];
  for (const node of folderNodes) {
    const idNode = (node as Element).getElementsByTagNameNS(NS.t, "FolderId")[0];
    const nameNode = (node as Element).getElementsByTagNameNS(NS.t, "DisplayName")[0];
    const countNode = (node as Element).getElementsByTagNameNS(NS.t, "TotalCount")[0];

    const id = idNode ? idNode.getAttribute("Id") || "" : "";
    const displayName = nameNode && nameNode.textContent ? nameNode.textContent : "";
    const totalCount = countNode && countNode.textContent ? parseInt(countNode.textContent, 10) : undefined;

    if (id && displayName && !EXCLUDED_FOLDER_NAMES.has(displayName)) {
      folders.push({ id, displayName, totalCount });
    }
  }
  return folders;
}

/** Moves the given item into the given folder. Throws on failure. */
export async function moveItemToFolder(itemId: string, folderId: string): Promise<void> {
  const xml = await makeEwsRequest(buildMoveItemRequest(itemId, folderId));
  const doc = parseXml(xml);
  assertNoEwsError(doc);
}
