# Tiller AI AutoCat
Apps Script code to use Gemini to automatically categorize financial transactions (designed to work with Tiller Finance Feeds and Google Sheets)

## About
- This is a script that is desined to work with the Tiller finance product to automatically categorize and clean up the Description column of your transactions (so you don't have to do it all manually!).
- It will only touch transactions that don't have a Category set.
- It works by trying to find how you've previously categorized transactions like the one it's working on, sending those to Gemini on Vertex AI, and asking it to do it's magic.  It will set the Category and Description field based on what comes back.
- It will pick the best valid category from your Category list, or fall back to a category you specify if it gets confused.
- There are no API keys to manage.  Calls to Vertex AI are authenticated with Application Default Credentials: `ScriptApp.getOAuthToken()` returns an OAuth token for whoever runs the script, and Vertex AI authorizes it against your Google Cloud project via IAM.
- If you want to mark transactions that have been modified by this code, add a column to your Transactions sheet called "AI AutoCat" - it will mark transactions it's modified by writing TRUE into this column.
- Given how sensitive this is to data, any and all feedback about how it's working (or not) is greatly appreciated.
- Special thanks to [@Aag1024](https://github.com/aag1024) for adding gemini suppport and the tfidf search module which works a lot better than my original hackery.

## Demo Video
- You can see this working with some sample data here: https://drive.google.com/file/d/16ROtqWboSOaNfgKGs0hUSjc3heGqFPBD/view?usp=drive_link

## Google Cloud Setup
You need a Google Cloud project with billing enabled.  Vertex AI usage is billed to that project.

1. Create (or pick) a project at https://console.cloud.google.com and note its **Project ID** and **Project Number**.
2. Enable the Vertex AI API for that project: https://console.cloud.google.com/apis/library/aiplatform.googleapis.com
3. Make sure the Google account that will run the script has the **Vertex AI User** (`roles/aiplatform.user`) role on the project.  If it's your own project and you're the owner, you already do.

## Installation Instructions
- From your Tiller connected Google Sheet, go to Extensions --> Apps Script
- If you don't have any existing Apps Script, you should just see Code.gs in the Files section on the left.
- Use the + button to add three new files called "gviz", "ai_autocat", and "tfidf_search".
- Copy and paste the contents of the files here into those files.
- Add (or change if you have one already) an OnOpen fuction to your code.gs file that matches the one here.  This just adds a menu item to call the AI AutoCat code.
- Link the Apps Script project to your Google Cloud project: in the Apps Script editor go to **Project Settings** (the gear icon) --> **Google Cloud Platform (GCP) Project** --> **Change project**, and enter your GCP **Project Number** (the numeric one, not the Project ID).  Without this step the OAuth token won't be accepted by Vertex AI.
- Still in **Project Settings**, tick **"Show appsscript.json manifest file in editor"**.  Open the `appsscript.json` that appears and add the `oauthScopes` array from the `appsscript.json` in this repo.  (Keep your own `timeZone` - you only need the scopes.)  The `cloud-platform` scope is the one that lets the script call Vertex AI; Apps Script will not ask for it on its own.
- Set your Google Cloud **Project ID** as a script property.  Still in **Project Settings**, scroll down to **Script Properties**, click **Add script property**, enter `GCP_PROJECT_ID` as the name and your project ID as the value, then **Save script properties**.  The script reads the project ID from here rather than from the code, so it stays out of version control - there is no project ID in any of the `.gs` files, and pasting in a new version of `ai_autocat.gs` will not disturb it.
- `GCP_LOCATION` at the top of ai_autocat.gs defaults to `global`, which routes to whichever region has capacity.  That gives the best availability and the widest model support, and is the right choice unless you need requests pinned to one geography.  If you do, set it to a specific region instead (`us-west1` is the west coast region with the broadest Gemini coverage; `us-central1` is the most widely supported overall).
- Modify ai_autocat.gs to use the FALLBACK_CATEGORY you want to use (this must be a valid category, or the empty string).
- Run `categorizeUncategorizedTransactions` once from the Apps Script editor and accept the authorization prompt.  You'll see a warning that the script wants to "See, edit, configure, and delete your Google Cloud data" - that's the `cloud-platform` scope, and it's required.

## Usage Instructions
- After installing the script, refresh your Tiller sheet.  You should see a new menu item appear called "AI AutoCat" after a few seconds.  You can run the AI autocat code manually from this menu item.
- If you want, you can also add a trigger to automatically run the AI AutoCat code nightly.  See instructions here: https://developers.google.com/apps-script/guides/triggers/installable.  The function you want to run is categorizeUncategorizedTransactions.
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

The baseline revision is pinned to the commit before this work landed, so the
comparison stays meaningful as further changes are committed. Override it with
`BASELINE_REV` to compare against a different revision:

```
BASELINE_REV=main npm test
```

## Troubleshooting
- **HTTP 403 with `PERMISSION_DENIED`** - either the Vertex AI API isn't enabled on the project, or the account running the script lacks `roles/aiplatform.user`.
- **HTTP 403 mentioning the caller's project, or 401** - the Apps Script project probably isn't linked to your GCP project number (Project Settings --> Change project), or the `cloud-platform` scope is missing from the manifest.  After changing scopes you have to re-run the script and re-accept the authorization prompt.
- **HTTP 404** - check `GEMINI_MODEL` is available in `GCP_LOCATION`.  Not every model is served from every region; `us-west2` and `us-west3` in particular do not serve Gemini models.  Switching `GCP_LOCATION` back to `global` is the quickest way to confirm the region is the problem.
