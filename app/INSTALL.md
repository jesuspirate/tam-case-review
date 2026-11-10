# TAM Case Review: install and use

Turns a ServiceNow case export into the page you present on the customer call.
Nothing else to install, no Terminal.

## Install

**Mac**

1. Open **TAM Case Review.dmg** and drag the app onto the **Applications** folder.
2. Open it from Applications. The first time, macOS says it can't verify the app.
   Click **Done**, then open **System Settings → Privacy & Security**, scroll down,
   and click **Open Anyway** next to "TAM Case Review". Confirm once. After that it opens normally.

**Windows**

1. Run **TAM Case Review Setup.exe**. No admin rights needed.
2. If Windows shows "Windows protected your PC", click **More info → Run anyway**.
3. Open it from the Start menu (or the desktop shortcut).

## Each review

1. **Export from ServiceNow** (the first time for each customer):
   1. Press **Create a new list**.
   2. Select a list you already have for your customer, or create one from scratch.
   3. Add a condition: **Customer account** is **Acme** (your customer).
   4. In **Select columns**, add the recommended columns (see the ServiceNow view guide,
      or click **Which columns?** in the app). Any columns work, but more columns give a
      better review.
   5. Press **Create**.
   6. In the list shown, press **Export**.

   After that, just open the saved list and press **Export**.
2. **Add it:** drag the downloaded file into the app (or click **Choose export file…**).
   Pick the account, or choose **New account…** and type a short name like `acme`.
   The date is filled in for you.
3. **Write updates:** the app shows one case at a time with its status, people, Jira
   ticket and the latest ServiceNow comment and internal note.
   - Type what the customer should hear, then click **Save & next** (or ⌘/Ctrl+Enter).
   - **Use latest public comment** or **Use this as my update** copies a ServiceNow comment instead.
   - **Skip** leaves the case as it is.
   - The list on the left shows **Needs update** (orange dot) and **updated today** (green dot).
     **All cases** shows the rest too.
4. **Build page:** the page opens in your browser. Present from there, or select all
   and copy it into OneNote. The Gainsight text is already on your clipboard, so paste it
   into the Gainsight activity note.
   The **Before the call** list is just for you. It's not on the page.

**Waiting on us** (home screen, named after your company) is an internal list of every case waiting on
us, across all your accounts.

## Where your data is

`Documents/TAM Case Review`, in three folders: `exports` (your ServiceNow files), `notes`
(your updates), and `out` (the pages). Settings → **Open folder** takes you there.
They contain customer case data, so keep them on your machine and don't share them.
