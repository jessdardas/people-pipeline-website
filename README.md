# People Pipeline – website (IIS)

The same People Pipeline app as the Google Apps Script version, as a normal website on your Windows **IIS** server.
It is a **separate project**: nothing here touches the Apps Script project, Google Drive or clasp.

## How it works

```
data\people pipeline.xlsx  ──►  the browser reads it (SheetJS)  ──►  server\Pipeline.js links the sheets  ──►  app\ draws the page
        ▲                                                             (same file as Apps Script)             (same folder as Apps Script)
        └── tools\copy-excel.ps1, every hour (Task Scheduler)
```

| Folder / file | What it is |
|---|---|
| `index.html` | start page: puts `app\index.html` together in the browser (like Apps Script does on its server) |
| `app\` | **copy** of the Apps Script project's `app\` folder: the page, its styles and code |
| `server\Pipeline.js` | **copy** of the Apps Script project's `server\Pipeline.js`: how the Excel sheets are read and linked |
| `site\source.js` | the website's only own code: reads the Excel file from `data\` and talks to `api\settings.ashx` |
| `site\vendor\` | SheetJS, xlsx-js-style, jsPDF (+ autotable): read Excel, export Excel / PDF, without internet |
| `data\people pipeline.xlsx` | **the data**: put the Excel file here (not in git) |
| `api\settings.ashx` | saves the **New DB** date for everyone after the admin password (needs ASP.NET, below) |
| `App_Data\settings.json` | the saved New DB date (made by the handler, not downloadable) |
| `web.config` | IIS settings: start page, `.json` / `.xlsx` allowed, no caching, hides `App_Data` and `tools` |
| `tools\copy-excel.ps1` | copies the newest Excel file into `data\` (run it every hour) |
| `tools\update-from-app.ps1` | copies the latest `app\` + `server\Pipeline.js` from the Apps Script project |

## Install on IIS (one time)

1. Copy this folder to the server, e.g. `C:\inetpub\people-pipeline`.
2. **IIS Manager** → Sites → *Add Website* (or *Add Application* under an existing site) → physical path = that folder.
3. **ASP.NET** (for saving the admin date only): *Server Manager → Add Roles and Features → Web Server (IIS) → Application Development → ASP.NET 4.8*. The app pool must be **.NET CLR v4**.
4. Give the app pool write access to `App_Data`: right-click `App_Data` → Properties → Security → Edit → Add `IIS AppPool\<your app pool name>` → **Modify**.
5. Put `people pipeline.xlsx` in `data\` and open the site.

Without steps 3–4 the site still works; only saving the New DB date doesn't (it then reads `data\settings.json`, which you can edit by hand).

## Fresh data every hour

1. Open `tools\copy-excel.ps1` and set `$Source` (where your query export saves the Excel file) and `$Target`.
2. **Task Scheduler** → *Create Task* → Triggers: *Daily*, repeat every **1 hour** → Actions: *Start a program*
   `powershell.exe` with arguments `-ExecutionPolicy Bypass -File "C:\inetpub\people-pipeline\tools\copy-excel.ps1"`.

The script also copies the **"last refreshed" notepad** (a `.txt` with *refresh* in its name, in the same folder as the Excel file) to `data\last refreshed.txt`. The top of the page shows its first line as written ("Last refreshed **28/09/2026 10:00**"); without that file it shows when the Excel file last changed.

Open pages read the files again every hour (and within 10 minutes of a new upload). There is no refresh button.

## Website-only features

These are only in the website's `app\` (not in the Apps Script project):

- **Map 01 · I.N. column**: open projects whose project name (or account name) has *I.N.*, *test* or *intercompany* go in the **I.N.** column instead of their phase column (`IN_NAME` in `app\core\config-js.html`).
- **Map 01 · "All projects ▾" menu** (next to the pipelines): tick *I.N. projects*, *Remove from pipeline* and/or *No P2* to take those projects off the map. An account whose open projects are all taken off leaves the map, so the totals go down. The two columns are found in the projects sheet by name (*remove … pipeline*, *no p2*); a cell counts when it is filled in and not *no* / *false* / *0* (`PROJECT_EXCLUDE` in `app\core\config-js.html`).
- **Account table**: *Phase* is hidden by default, like *Industry* (show it again with **Columns**).
- **Account details**: click any column title in the tables (projects, contact persons, social media…) to sort A→Z / low→high, click again for Z→A / high→low.
- **Last refreshed** comes from the notepad (see above); no ↻ button.

## Updating the website after changing the app

The website uses the same page code as the Apps Script project. After you change something there:

```powershell
powershell -ExecutionPolicy Bypass -File tools\update-from-app.ps1 -AppProject "C:\Users\<you>\people pipeline"
```

then copy the folder to the server again (or run the script on the server). Only `app\` and `server\Pipeline.js` are copied.

⚠️ This **replaces** `app\`, so it removes the website-only features above until they are also added to the Apps Script project.

## Admin

Same as the Apps Script app: the tiny dot at the bottom right → password → New DB date (saved for everyone) and Data check.
The password is not in the code, only its fingerprint (`AdminHash` in `api\settings.ashx`, the same as `ADMIN_HASH` in the Apps Script project).
