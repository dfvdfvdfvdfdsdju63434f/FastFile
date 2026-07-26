# Smart Filing — Outlook Add-in

Suggests a folder to file the currently open email into, based on the
subject, sender, and a short preview of the body — entirely on-device, with
no backend and no external API calls.

## How it decides

This build implements the scope below (matching the answers used to design
it — see **Roadmap** for the deferred parts of the original spec):

| Decision | This build |
|---|---|
| Suggestion logic | Local, rule-based keyword/sender matching (`src/services/folder-suggester.service.ts`) |
| Filing scope | The currently open email only (no thread-wide filing) |
| Backend | None — everything runs in the task pane |

Because there's no LLM call and no backend, folder suggestions and the move
itself both happen through **`Office.context.mailbox.makeEwsRequestAsync`**
— Office.js's built-in bridge to Exchange Web Services. It requires no Azure
AD app registration, no OAuth consent screen, and no server to run — Outlook
signs and scopes every request to the current user's own mailbox.

## Project structure

```
FastFile/
├── manifest.xml                    # Add-in manifest (XML, AppSource-style)
├── package.json
├── tsconfig.json
├── webpack.config.js
├── assets/                         # Ribbon/store icons (generated PNGs)
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   └── icon-128.png
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html           # Task pane markup
│   │   ├── taskpane.css            # Fluent-inspired styling (no CDN dependency)
│   │   └── taskpane.ts             # Office.onReady lifecycle + UI logic
│   ├── commands/
│   │   ├── commands.html           # Required "function file" host page
│   │   └── commands.ts             # Minimal Office.actions.associate({}) boilerplate
│   └── services/
│       ├── office-item.service.ts  # Reads subject/sender/body from the open item
│       ├── ews.service.ts          # EWS calls: list folders, move item
│       └── folder-suggester.service.ts  # Rule-based scoring engine
└── .github/workflows/deploy.yml    # Builds + publishes dist/ to GitHub Pages
```

## Prerequisites

- Node.js 18+ and npm
- A Microsoft 365 / Exchange Online mailbox (or Exchange on-prem with EWS enabled)
- **Outlook desktop (classic, Windows or Mac) or Outlook on the web** for testing
  — see the EWS limitation note below before testing on "new Outlook" or mobile
- Git and a GitHub account, for hosting

## 1. Install and build

```bash
npm install
npm run build        # outputs to dist/
```

`npm run build` bundles `taskpane.ts` and `commands.ts` with webpack and
copies the HTML/CSS/assets/manifest alongside them into `dist/`, ready to be
hosted as static files.

## 2. Host it

### Option A — GitHub Pages (recommended for this project)

1. Push this project to a GitHub repository.
2. In the repo, go to **Settings → Pages** and set the source to
   **GitHub Actions**.
3. The included workflow (`.github/workflows/deploy.yml`) builds and deploys
   `dist/` to Pages automatically on every push to `main`. You can also
   trigger it manually from the **Actions** tab (`workflow_dispatch`).
4. Once deployed, note your Pages URL — it will look like
   `https://<your-username>.github.io/<repo-name>/`.
5. **Update `manifest.xml`**: replace every occurrence of
   `https://your-username.github.io/FastFile` with your actual
   Pages URL (there are several: `IconUrl`, `HighResolutionIconUrl`,
   `SupportUrl`, `AppDomains`, `SourceLocation`, and the three `bt:Url` /
   `bt:Image` entries under `Resources`). Commit and push again so the
   deployed `manifest.xml` matches.
6. Validate it: `npm run validate` (uses `office-addin-manifest`).

### Option B — Local dev server (fast iteration, before you push)

```bash
npm run dev-server
```

This serves `dist/` over HTTPS at `https://localhost:3000` using a
self-signed cert (installed automatically the first time via
`office-addin-dev-certs`, which `office-addin-debugging` calls under the
hood). To sideload against this, temporarily point the manifest URLs at
`https://localhost:3000/...` instead of your GitHub Pages URL, or just use
`npm start` below, which handles both build and sideload for you.

## 3. Sideload into Outlook

**Fastest path (desktop, automated):**

```bash
npm start
```

This runs `office-addin-debugging start`, which builds the project, starts
the dev server, and sideloads the manifest into Outlook desktop
automatically. Run `npm run stop` to remove it afterward.

**Manual sideload (any platform, using your GitHub Pages manifest):**

- **Outlook on the web / new Outlook:**
  Settings (gear icon) → **Mail → Customize actions → Add-ins**, or go
  directly to **Get Add-ins → My add-ins → Add a custom add-in → Add from
  file**, and select your `manifest.xml`.
- **Outlook desktop (Windows), classic:**
  **File → Manage Add-ins** (opens OWA add-in management) and sideload as
  above — sideloaded add-ins sync to desktop Outlook shortly after.
- **Outlook desktop (Mac):**
  **Tools → Get Add-ins → My add-ins → Add a custom add-in → Add from file**.

## 4. Test it

1. Open any email in the reading pane (or in its own window).
2. On the ribbon, look for the **Smart Filing** group and click
   **Suggest Folder**.
3. The task pane opens, reads the subject/sender/body preview, fetches your
   mail folders, and shows up to 4 ranked suggestions with a plain-English
   reason for each (e.g. *"Subject mentions invoice, payment"*).
4. The top suggestion is pre-selected. Click a different card to change the
   selection, or use **Choose manually** to pick from the full folder list.
5. Click **File email**. On success you'll see a confirmation banner; the
   message is moved via EWS `MoveItem`.

### Things worth testing deliberately

- An email from a sender whose domain matches an existing folder name.
- An email with no good keyword match — confirm the empty state and manual
  picker both work.
- Clicking **File email** twice in a row (the button disables mid-request).
- Revoking network/EWS in a client that doesn't support it (see below) —
  confirm the error banner is legible, not a stack trace.

## Known limitation: EWS availability

`makeEwsRequestAsync` — and therefore this add-in's folder listing and
move — depends on the Outlook client supporting EWS. It works reliably on:

- Outlook desktop (classic Windows/Mac)
- Outlook on the web (classic)

It is **not guaranteed** on:

- Some "new Outlook for Windows" builds and Outlook for Mac's newer UI, as
  Microsoft migrates these toward Graph-only APIs
- Outlook mobile (iOS/Android)

The add-in detects this and shows a clear error (`"This Outlook client
doesn't support direct folder access..."`) rather than failing silently —
but if your users are primarily on those clients, plan for the Graph-based
upgrade path below sooner rather than later.

## Roadmap: extending beyond this build

The original spec included two things intentionally deferred here:

**1. Filing the entire thread, not just the open email.**
Office.js only gives read access to the *currently open* item — moving
other messages in the same conversation requires the Microsoft Graph API
(`GET /me/messages?$filter=conversationId eq '...'` followed by batched
`POST /move` calls), which needs:
- An Azure AD app registration
- Office.js SSO (`OfficeRuntime.auth.getAccessToken`) to get a token for the
  signed-in user
- The `Mail.ReadWrite` Graph delegated permission (and admin consent if
  deployed org-wide)

**2. LLM-powered suggestions instead of keyword rules.**
Swap `folder-suggester.service.ts`'s `suggestFolders()` for a call to an LLM
endpoint. Do **not** call the LLM API directly from the task pane with an
embedded API key — put a thin backend in front of it (even a single Azure
Function / Cloudflare Worker) that:
- Accepts `{ subject, senderName, senderEmail, bodyPreview, folderNames }`
- Holds the LLM API key server-side
- Optionally uses Office.js SSO to authenticate the caller
- Returns a ranked folder name + confidence, which you map back to a
  `MailFolder` client-side

Both upgrades are additive — the EWS move logic and the UI shell here don't
need to change, only the suggestion source and (for thread-filing) the
"file all in thread" checkbox and its Graph-backed move loop.

## Permissions note

The manifest requests `ReadWriteMailbox`, the highest Office.js permission
level. This is required for `makeEwsRequestAsync` to work at all — Office.js
doesn't offer a narrower scope that still allows folder listing + moving
items. It does **not** grant access to other users' mailboxes or anything
outside the signed-in user's own mailbox.
