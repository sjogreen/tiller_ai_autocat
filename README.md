# Tiller AI AutoCat
Apps Script code to use Gemini to automatically categorize financial transactions (designed to work with Tiller Finance Feeds and Google Sheets)

## About
- This is a script that is desined to work with the Tiller finance product to automatically categorize and clean up the Description column of your transactions (so you don't have to do it all manually!).
- It will only touch transactions that don't have a Category set.
- It works by trying to find how you've previously categorized transactions like the one it's working on, sending those to Gemini on Vertex AI, and asking it to do it's magic.  It will set the Category and Description field based on what comes back.  See "How it works" below for the details, including exactly what is sent where.
- It will pick the best valid category from your Category list, or fall back to a category you specify if it gets confused.
- There are no API keys to manage for Gemini (the optional Jev stage needs an OpenRouter key).  Calls to Vertex AI are authenticated with Application Default Credentials: `ScriptApp.getOAuthToken()` returns an OAuth token for whoever runs the script, and Vertex AI authorizes it against your Google Cloud project via IAM.
- If you want to mark transactions that have been modified by this code, add a column to your Transactions sheet called "AI AutoCat" - it will mark transactions it's modified by writing TRUE into this column.
- Given how sensitive this is to data, any and all feedback about how it's working (or not) is greatly appreciated.
- Special thanks to [@Aag1024](https://github.com/aag1024) for adding gemini suppport and the tfidf search module which works a lot better than my original hackery.

## How it works

Each run takes up to 50 transactions that have a Full Description and no Category.
Any of those that have no Transaction ID (usually rows added by hand) are first
given one, `autocat:` followed by a [ULID](https://github.com/ulid/spec), so the
answer can be written back to the right row.  Only rows the run is about to send
get an ID, and only their Transaction ID cell is written.

1. **Find previous transactions.** For each one, a TF-IDF word-overlap search
   over the Full Description of your categorized transactions from the last
   365 days finds up to 6 similar ones.  When more than 6 of the top 20 matches
   tie exactly on word overlap (every "CHECK #1234" looks alike), the ones
   closest in amount are chosen instead, the closest of each category first.
   This runs inside the sheet and sends nothing anywhere.
2. **Jev (optional).** If the `OPENROUTER_API_KEY` script property is set, each
   transaction that has previous transactions is sent to TypeSafe's Jev
   (`typesafe/jev-1.13`) through OpenRouter, 8 at a time.  Jev picks the previous
   transaction that is the same merchant or recurring payment, or "none".  When
   it picks one with at most a 50% chance of "none", that transaction's
   Description and Category are copied and Gemini is not asked about it.  A pick
   whose category is your FALLBACK_CATEGORY is ignored.
3. **Gemini.** Everything Jev did not settle (or everything, without the key) goes to
   Gemini on Vertex AI in one request, with temperature 0, minimal thinking and
   a fixed JSON response schema.  It returns a cleaned description and a category
   for each, or declines, which is written as FALLBACK_CATEGORY.

This is the same pipeline, prompt and settings as the Compound categorizer, except
that categories are sent by name rather than id and there is no bank
transaction-type hint.

**What is sent.**  Only these fields, and nothing else from your sheet (no account
names, balances or notes):

| | Jev, via OpenRouter (only with the key) | Gemini, via Vertex AI |
|---|---|---|
| Transaction | Full Description, Amount, Date | Transaction ID, Full Description, Amount, Date |
| Each previous transaction | Full Description, Category, Amount, Date | Full Description, Description, Category, Amount, Date |
| Categories | - | Every category name, with its Group when the Categories sheet has a Group column |

The Jev stage is off until you add an OpenRouter key; see "Optional: the Jev
stage" under Configuring it.

## Demo Video
- You can see this working with some sample data here: https://drive.google.com/file/d/16ROtqWboSOaNfgKGs0hUSjc3heGqFPBD/view?usp=drive_link

## Google Cloud Setup
You need a Google Cloud project with billing enabled.  Vertex AI usage is billed to that project.

1. Create (or pick) a project at https://console.cloud.google.com and note its **Project ID** and **Project Number**.
2. Enable the Vertex AI API for that project: https://console.cloud.google.com/apis/library/aiplatform.googleapis.com
3. Make sure the Google account that will run the script has the **Vertex AI User** (`roles/aiplatform.user`) role on the project.  If it's your own project and you're the owner, you already do.

## Installation Instructions

### Getting the code into your sheet

There are two ways to do this.  Copy and paste is fine for a one-off install;
`clasp` is worth it if you expect to pick up later versions.

**Copy and paste**

- From your Tiller connected Google Sheet, go to Extensions --> Apps Script
- If you don't have any existing Apps Script, you should just see Code.gs in the Files section on the left.
- Use the + button to add three new files called "gviz", "ai_autocat", and "tfidf_search".  Leave off the `.gs` - Apps Script adds it for you, and typing it produces a file called `tfidf_search.gs.gs`.
- Copy and paste the contents of the files here into those files.
- Add (or change if you have one already) an onOpen function to your code.gs file that matches the one here.  This just adds the menu items that call into this code.

**With clasp**

[clasp](https://github.com/google/clasp) pushes the files straight into the
script project, and lets you diff your sheet against this repo later to see
whether it has fallen behind.

- Turn on the Apps Script API for your account (once): https://script.google.com/home/usersettings
- Clone this repo, then `npm install` and `npx clasp login`
- Get the script ID from your sheet: Extensions --> Apps Script --> Project Settings (gear icon) --> Script ID
- Write it into a `.clasp.json` beside package.json.  This file is gitignored, since it names one specific spreadsheet:

```json
{ "scriptId": "YOUR_SCRIPT_ID_HERE", "rootDir": "." }
```

- `npm run sheet:status` lists what would be pushed - it should be exactly `appsscript.json` and the four `.gs` files
- `npm run sheet:push` uploads them

To check an existing install against this repo, `npm run sheet:pull` and then
`git diff`.  A clean diff means the sheet is current.

**`clasp push` replaces the script project's entire file set.**  Anything in the
project that is not in your local directory is removed.  If your sheet has Apps
Script code of its own, read "Keeping your own code alongside this" below before
your first push.

### Configuring it

- Link the Apps Script project to your Google Cloud project: in the Apps Script editor go to **Project Settings** (the gear icon) --> **Google Cloud Platform (GCP) Project** --> **Change project**, and enter your GCP **Project Number** (the numeric one, not the Project ID).  Without this step the OAuth token won't be accepted by Vertex AI.
- Still in **Project Settings**, tick **"Show appsscript.json manifest file in editor"**.  Open the `appsscript.json` that appears and add the `oauthScopes` array from the `appsscript.json` in this repo.  (Keep your own `timeZone` - you only need the scopes.)  The `cloud-platform` scope is the one that lets the script call Vertex AI; Apps Script will not ask for it on its own.
- Set your Google Cloud **Project ID** as a script property.  Still in **Project Settings**, scroll down to **Script Properties**, click **Add script property**, enter `GCP_PROJECT_ID` as the name and your project ID as the value, then **Save script properties**.  The script reads the project ID from here rather than from the code, so it stays out of version control - there is no project ID in any of the `.gs` files, and pasting in a new version of `ai_autocat.gs` will not disturb it.
- `GCP_LOCATION` at the top of ai_autocat.gs defaults to `global`, which routes to whichever region has capacity.  That gives the best availability and the widest model support, and is the right choice unless you need requests pinned to one geography.  If you do, set it to a specific region instead (`us-west1` is the west coast region with the broadest Gemini coverage; `us-central1` is the most widely supported overall).
- Modify ai_autocat.gs to use the FALLBACK_CATEGORY you want to use (this must be a valid category, or the empty string).
- Run `categorizeUncategorizedTransactions` once from the Apps Script editor and accept the authorization prompt.  You'll see a warning that the script wants to "See, edit, configure, and delete your Google Cloud data" - that's the `cloud-platform` scope, and it's required.

### Where settings and keys are kept

Nothing secret, and nothing specific to your sheet, is ever in the `.gs` files.
Everything of that kind lives in **Script Properties** (Apps Script editor -->
Project Settings --> Script Properties), which belong to your script project in
Google's cloud.  They are not part of any file, so they are never committed to
git, and `clasp push`/`clasp pull` and copy-paste upgrades leave them alone.

| Script property | Required | What it is |
|---|---|---|
| `GCP_PROJECT_ID` | Yes | Your Google Cloud project ID. Not a secret - Vertex AI authorizes your Google sign-in, not a key |
| `OPENROUTER_API_KEY` | No | An OpenRouter API key. Turns on the Jev stage |

Anyone with edit access to the script project can read Script Properties, so
use a key you would be comfortable sharing with your sheet's editors.

### Optional: the Jev stage

Jev is a model from TypeSafe that decides whether a transaction is the same
merchant as one you have already categorized, and if so copies that category and
description.  It is cheap (about $0.00004 per transaction) and settles most
recurring transactions without asking Gemini.  It is reached through OpenRouter,
so turning it on sends the fields listed in "How it works" to OpenRouter and
TypeSafe as well as to Google.

1. Create an account at https://openrouter.ai, add credit, and create an API key
   under **Keys**.  Setting a credit limit on the key is a good idea.
2. In the Apps Script editor go to **Project Settings** --> **Script Properties**
   --> **Add script property**.  Enter `OPENROUTER_API_KEY` as the name and the
   key as the value, then **Save script properties**.
3. Run AutoCat.  The log now shows a `jevAsked` / `jevTaken` / `jevFailed` /
   `jevCost` summary and a "Jev decisions" line for each transaction it was asked
   about.

To turn it off, delete the property; the script goes back to Gemini only.

## Keeping your own code alongside this

Apps Script allows only one `onOpen` per project, so a sheet that already has its
own menu code cannot simply take this repo's `code.gs`.  Rather than hand-merging
it (which makes `code.gs` impossible to update), define `addLocalMenus` in a file
of your own and `onOpen` will call it if it exists, passing the `Ui`:

```javascript
// personal.gs - your file, not part of this repo
function addLocalMenus(ui) {
  ui.createMenu('My Own Menu')
      .addItem('My Own Thing', 'myOwnFunction')
      .addToUi();
}

function myOwnFunction() { ... }
```

Your items go in a menu of your own, separate from "Tiller AI AutoCat".

If you install with clasp, keep that file in the repo directory so it gets pushed
with the rest, and add it to `.gitignore` so it never gets committed.  `personal.gs`
is already gitignored for this purpose.

## Usage Instructions
- After installing the script, refresh your Tiller sheet.  You should see a new menu appear called "Tiller AI AutoCat" after a few seconds.  It has two items: **Run AutoCat**, which categorizes your uncategorized transactions, and **Search Transactions (Active Cell)**, which shows the previously categorized transactions that look most like the one your cursor is on.
- If you want, you can also add a trigger to automatically run the AI AutoCat code nightly.  See instructions here: https://developers.google.com/apps-script/guides/triggers/installable.  The function you want to run is categorizeUncategorizedTransactions.
- Each run logs what it did: the transactions and previous transactions it found, Jev's decision on each (when Jev is on), Gemini's answers, and what it wrote.  Open them from the Apps Script editor's **Executions** page, or with `npx clasp logs` once `.clasp.json` has your project ID (see Development).
- Each run logs an estimated cost.  That estimate uses the `INPUT_COST_PER_M_TOKENS` / `OUTPUT_COST_PER_M_TOKENS` constants at the top of ai_autocat.gs - update them if you change `GEMINI_MODEL` or if pricing changes.

## Development

The `.gs` files run in Apps Script, but the parts that do not touch the Sheets UI
are covered by tests that run under Node (no dependencies, Node 18+):

```
npm test
```

The tests load each `.gs` file into a sandbox with the Apps Script globals stubbed
out, and compare the current behaviour against the previously committed version of
the same file. They exist mainly to protect two invariants that are easy to break:

- **Columns can appear in any order.** Every column is resolved from its header
  name, never from a fixed position.
- **Only the specific cells being changed are ever written.** Tiller sheets often
  drive an entire column from a single ARRAYFORMULA, and writing a literal into
  such a column destroys it. `planCellWrites` batches writes into rectangular
  blocks only where every cell in the block is one the script meant to write, and
  the tests assert that the planned cells match the intended cells exactly.

If you have a `.clasp.json` set up (see the install instructions), these push to
and pull from the sheet its script ID points at:

```
npm run sheet:status   # list the files clasp would push
npm run sheet:pull     # fetch the script project's current source
npm run sheet:push     # upload the local .gs files to the script project
```

`.claspignore` denies everything and then re-allows only `appsscript.json` and the
`.gs` files, so the Node tests are never uploaded into the script project.
`sheet:push` passes `--force` because clasp otherwise prompts before overwriting
the remote manifest, and skips the whole push when it cannot read an answer.

To read run logs from the command line, add your GCP project ID to `.clasp.json`
(it is gitignored, like the rest of that file):

```json
{ "scriptId": "YOUR_SCRIPT_ID_HERE", "rootDir": ".", "projectId": "your-gcp-project-id" }
```

Then `npx clasp logs` prints the most recent log entries from Cloud Logging.

The baseline revision is pinned to the commit before this work landed, so the
comparison stays meaningful as further changes are committed. Override it with
`BASELINE_REV` to compare against a different revision:

```
BASELINE_REV=main npm test
```

## Troubleshooting
- **HTTP 403 with `PERMISSION_DENIED`** - either the Vertex AI API isn't enabled on the project, or the account running the script lacks `roles/aiplatform.user`.
- **HTTP 403 mentioning the caller's project, or 401** - the Apps Script project probably isn't linked to your GCP project number (Project Settings --> Change project), or the `cloud-platform` scope is missing from the manifest.  After changing scopes you have to re-run the script and re-accept the authorization prompt.
- **HTTP 400 "Thinking level is unsupported"** - the model does not accept `GEMINI_THINKING_LEVEL`.  `gemini-3.8-flash` on Vertex AI rejects `MINIMAL`; the script uses `LOW`.  If you change `GEMINI_MODEL`, check which levels it supports.
- **`jevFailed` above 0** - look for "Jev request failed" lines in the log.  HTTP 401 means `OPENROUTER_API_KEY` is wrong or revoked; 402 means the OpenRouter account is out of credit.  Failed transactions still go to Gemini, so the run completes either way.
- **HTTP 404** - check `GEMINI_MODEL` is available in `GCP_LOCATION`.  Not every model is served from every region; `us-west2` and `us-west3` in particular do not serve Gemini models.  Switching `GCP_LOCATION` back to `global` is the quickest way to confirm the region is the problem.
