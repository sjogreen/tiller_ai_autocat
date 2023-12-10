# Tiller AI AutoCat
Apps Script code to use Open AI to automatically categorize financial transactions (designed to work with Tiller Finance Feeds and Google Sheets)

## About
- This is a script that is desined to work with the Tiller finance product to automatically categorize and clean up the Description column of your transactions (so you don't have to do it all manually!).
- It will only touch transactions that don't have a Category set.
- It works by trying to find how you've previously categorized transactions like the one it's working on, sending those to Open AI, and asking it to do it's magic.  It will set the Category and Description field based on what comes back.
- It will pick the best valid category from your Category list, or fall back to a category you specify if it gets confused.
- If you want to mark transactions that have been modified by this code, add a column to your Transactions sheet called "AI AutoCat" - it will mark transactions it's modified by writing TRUE into this column.
- This works for me, and I've tried to make it somewhat generic so it works for others -- but I DISCLAIM ALL RESPONSBILITY IF IT MESSES ANYTHING UP IN YOUR SHEET.  You can always undo or revert to a previous version.
- Given how sensitive this is to data, any and all feedback about how it's working (or not) is greatly appreciated.

## Installation Instructions
- First, you need to get an Open AI API Key to use.  Sign up as a developer with Open AI and get a secret key.
- From your Tiller connected Google Sheet, go to Extensions --> Apps Script
- If you don't have any existing Apps Script, you should just see Code.gs in the Files section on the left.
- Use the + button to add two new files called "gviz.gs" and "ai_autocat.gs".
- Copy and paste the contents of the files here into those files.
- Add (or change if you have one already) an OnOpen fuction to your code.gs file that matches the one here.  This just adds a menu item to call the AI AutoCat code.
- Modify ai_autocat.gs to use your Open AI API Secret Key
- Modify ai_autocat.gs to use the FALLBACK_CATEGORY you want to use (this must be a valid category, or the empty string).

## Usage Instructions
- After installing the script, refresh your Tiller sheet.  You should see a new menu item appear called "AI AutoCat" after a few seconds.  You can run the AI autocat code manually from this menu item.
- If you want, you can also add a trigger to automatically run the AI AutoCat code nightly.  See instructions here: https://developers.google.com/apps-script/guides/triggers/installable.  The function you want to run is categorizeUncategorizedTransactions.
