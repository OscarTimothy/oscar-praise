# Wedding expenses

A password protected page for the Praise & Oscar wedding budget. It shows the
dashboard (totals, spending by category, event and month, open balances and
every payment) and lets you add or edit expenses. Everything is read from and
written to the **Praise & Oscar Wedding Expenses** Google Sheet.

Once GitHub Pages is on for this repo, the page lives at
`https://oscartimothy.github.io/oscar-praise/expenses/`.

## How it works

```
GitHub Pages (expenses/index.html)  ──POST──▶  Apps Script web app  ──▶  Google Sheet
         password typed by you               checks the password,
                                              reads and writes rows
```

* The website holds **no data and no password**. It sends what you type to
  the Apps Script, which compares it with a password stored in the script's
  own settings. A wrong password gets nothing back, and 10 wrong tries in
  10 minutes locks it for 10 minutes.
* The script only writes the columns you fill in (A to I, L and M on the
  Expenses tab). It never touches the **Paid (USD equiv.)** and **Balance Due**
  formula columns.
* New expenses go on the next empty row. Before saving an edit, the script
  checks the row still has the same vendor and item, so two people editing
  at once can't overwrite each other.

## One time setup (about 5 minutes)

### 1. Add the script to the sheet

1. Open the **Praise & Oscar Wedding Expenses** sheet.
2. Go to **Extensions → Apps Script**.
3. Delete what's in `Code.gs` and paste in everything from
   [`apps-script/Code.gs`](apps-script/Code.gs). Save.

### 2. Set the password

1. In Apps Script, click the gear icon (**Project Settings**).
2. Under **Script Properties**, click **Add script property**.
3. Property: `PASSWORD`. Value: the password you want people to type. Save.

To change the password later, edit this value. No code changes are needed.

### 3. Deploy it

1. Click **Deploy → New deployment**.
2. Select type: **Web app**.
3. Execute as: **Me**. Who has access: **Anyone**.
4. Click **Deploy**, approve the Google permissions it asks for, and copy the
   **Web app URL** (it ends in `/exec`).

"Anyone" means anyone with the URL can reach the script, but every request
still needs the password.

### 4. Connect the site

Paste the URL into [`config.js`](config.js):

```js
window.EXPENSES_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/XXXX/exec',
};
```

Commit and push. Then make sure GitHub Pages is on: repo **Settings → Pages →
Deploy from a branch → main / (root)**.

## Updating the script later

After editing `Code.gs` in Apps Script, use **Deploy → Manage deployments →
Edit (pencil) → Version: New version → Deploy**. This keeps the same URL, so
`config.js` doesn't change.

## Files

| File | What it is |
|---|---|
| `index.html` | The whole site: password screen, dashboard, add and edit form |
| `config.js` | The Apps Script URL |
| `apps-script/Code.gs` | The backend you paste into the sheet's Apps Script |
