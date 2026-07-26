/**
 * Local, rule-based folder suggestion. No network calls, no LLM — everything
 * runs against data already on the client. Scoring is deliberately simple and
 * transparent (see `reasons`) so the UI can explain *why* a folder was
 * suggested, which matters for user trust in a "move my mail" feature.
 */

import { MailFolder } from "./ews.service";
import { EmailContext } from "./office-item.service";

export interface FolderSuggestion {
  folder: MailFolder;
  score: number; // relative score, higher = better match
  reasons: string[];
}

// Topic keyword groups: the group name itself is also matched against folder
// names (e.g. a folder called "Finance" or "Invoices" both benefit).
const KEYWORD_GROUPS: Record<string, string[]> = {
  finance: ["invoice", "receipt", "payment", "billing", "expense", "purchase", "order", "refund"],
  meetings: ["meeting", "agenda", "calendar", "schedule", "invite", "sync", "standup"],
  hr: ["payroll", "benefits", "onboarding", "timesheet", "leave", "vacation", "pto"],
  legal: ["contract", "agreement", "nda", "terms", "compliance", "policy"],
  projects: ["project", "sprint", "release", "deadline", "milestone", "status"],
  travel: ["flight", "itinerary", "hotel", "booking", "travel", "reservation"],
  support: ["ticket", "issue", "bug", "support", "incident", "case"],
  newsletters: ["newsletter", "digest", "subscription", "unsubscribe"],
};

const STOPWORDS = new Set([
  "re", "fw", "fwd", "the", "a", "an", "and", "or", "to", "of", "for", "on",
  "in", "is", "your", "you", "our", "com", "www", "with", "this", "that",
]);

function tokenize(text: string): string[] {
  return text
    .toLowerCase()
    .replace(/[^a-z0-9@.\s-]/g, " ")
    .split(/\s+/)
    .filter((t) => t.length > 1 && !STOPWORDS.has(t));
}

/**
 * Scores every folder against the email and returns the top matches,
 * highest score first. Folders with a score of 0 are omitted.
 */
export function suggestFolders(
  context: EmailContext,
  folders: MailFolder[],
  maxSuggestions = 4
): FolderSuggestion[] {
  const subjectTokens = tokenize(context.subject);
  const senderTokens = tokenize(`${context.senderName} ${context.senderEmail}`);
  const bodyTokens = tokenize(context.bodyPreview);
  const senderDomain = (context.senderEmail.split("@")[1] || "").split(".")[0];

  const scored: FolderSuggestion[] = folders.map((folder) => {
    const folderTokens = tokenize(folder.displayName);
    let score = 0;
    const reasons: string[] = [];

    // 1. Folder name words appear directly in the subject.
    const subjectOverlap = folderTokens.filter((t) => subjectTokens.indexOf(t) !== -1);
    if (subjectOverlap.length > 0) {
      score += subjectOverlap.length * 30;
      reasons.push(`Subject mentions "${subjectOverlap.join(", ")}"`);
    }

    // 2. Folder name words appear in the sender's name/address.
    const senderOverlap = folderTokens.filter((t) => senderTokens.indexOf(t) !== -1);
    if (senderOverlap.length > 0) {
      score += senderOverlap.length * 35;
      reasons.push("Sender matches this folder's name");
    }

    // 3. Sender's domain literally appears in the folder name
    //    (e.g. folder "Acme Corp" <- sender @acme.com).
    if (senderDomain && senderDomain.length > 2 && folder.displayName.toLowerCase().indexOf(senderDomain) !== -1) {
      score += 40;
      reasons.push(`Sender domain "${senderDomain}" matches this folder`);
    }

    // 4. Topic keyword groups: does the folder name relate to a known topic,
    //    and does the subject/body contain words from that topic?
    Object.keys(KEYWORD_GROUPS).forEach((group) => {
      const keywords = KEYWORD_GROUPS[group];
      const folderMatchesGroup = folderTokens.some(
        (t) => group.indexOf(t) !== -1 || t.indexOf(group.slice(0, 5)) !== -1
      );
      if (!folderMatchesGroup) return;

      const subjectHits = keywords.filter((k) => subjectTokens.some((t) => t.indexOf(k) !== -1));
      const bodyHits = keywords.filter((k) => bodyTokens.some((t) => t.indexOf(k) !== -1));

      if (subjectHits.length > 0) {
        score += subjectHits.length * 20;
        reasons.push(`Subject relates to "${group}" (${subjectHits.join(", ")})`);
      }
      if (bodyHits.length > 0) {
        score += bodyHits.length * 8;
        reasons.push(`Body mentions "${group}" topics`);
      }
    });

    // 5. Small tie-breaking boost for folders that are actively used.
    if (folder.totalCount && folder.totalCount > 0) {
      score += Math.min(5, Math.log10(folder.totalCount + 1));
    }

    return { folder, score: Math.round(score), reasons: dedupe(reasons) };
  });

  return scored
    .filter((s) => s.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, maxSuggestions);
}

function dedupe(items: string[]): string[] {
  return items.filter((item, index) => items.indexOf(item) === index);
}
