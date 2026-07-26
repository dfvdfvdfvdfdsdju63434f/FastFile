/**
 * Wraps Office.context.mailbox.item access behind promise-based, error-handled
 * helpers so the rest of the app never touches the raw callback APIs directly.
 */

export interface EmailContext {
  itemId: string;
  subject: string;
  senderName: string;
  senderEmail: string;
  bodyPreview: string;
  conversationId?: string;
}

const BODY_PREVIEW_MAX_CHARS = 600;

/**
 * Reads the fields we need from the currently open message.
 * Throws a descriptive Error if the item, its id, or its body can't be read —
 * callers should catch this and surface err.message to the user.
 */
export async function getCurrentEmailContext(): Promise<EmailContext> {
  const item = Office.context.mailbox && Office.context.mailbox.item;

  if (!item) {
    throw new Error("No email is currently open. Open a message and try again.");
  }

  const readItem = item as Office.MessageRead;

  const itemId = readItem.itemId;
  if (!itemId) {
    throw new Error(
      "This message doesn't have a saved item ID yet (it may be a brand-new draft). Try again after it syncs."
    );
  }

  const subject = readItem.subject || "(no subject)";
  const from = readItem.from;
  const senderName = (from && from.displayName) || "Unknown sender";
  const senderEmail = (from && from.emailAddress) || "";
  const conversationId = readItem.conversationId || undefined;

  const bodyPreview = await getBodyPreview(readItem, BODY_PREVIEW_MAX_CHARS);

  return { itemId, subject, senderName, senderEmail, bodyPreview, conversationId };
}

function getBodyPreview(item: Office.MessageRead, maxLength: number): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!item.body) {
      resolve("");
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result: Office.AsyncResult<string>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = result.value || "";
        resolve(text.substring(0, maxLength));
      } else {
        const message = (result.error && result.error.message) || "Failed to read the email body.";
        reject(new Error(message));
      }
    });
  });
}
