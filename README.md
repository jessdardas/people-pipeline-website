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

**Last updated at** (top of the page) is the last date written in `data\refresh-log.txt` (`C:\inetpub\people-pipeline\data\refresh-log.txt`), shown as written. Without that file it shows when the Excel file last changed.

Open pages read the files again every hour (and within 10 minutes of a new upload). There is no refresh button.

## Website-only features

These are only in the website's `app\` (not in the Apps Script project):

- **"Real projects only"** – a small switch above maps 01 and 02, **on** by default. It leaves out *I.N.*, *test* and *intercompany* projects (by project or account name), and projects marked *no P2* or *excluded from pipeline* (columns in the projects sheet). The small **▾** next to it picks which of the 5 kinds are left out (e.g. untick *Test* to see the tests). An account whose projects are all left out leaves the map, so the totals go down. Switch it off to see everything. On map 01, the I.N. / test / intercompany projects that are shown go in the **Internal / test** column. Settings: `PROJECT_EXCLUDE` in `app\core\config-js.html`.
- **Last project**: when the file has no last-project date for an account, the latest date of its own projects is used (map 02 and the details panel, which then says *from its projects*). The project date column is the first one found by `PROJECT_DATE_COLS` in `app\core\config-js.html`.
- **New DB**: with New DB on, a small date box next to *validated on or after* changes the date for your view only; nothing is saved (*reset* goes back to the admin date).
- **Account table**: *Last contact* next to *Last meeting*; *Phase* and *Industry* hidden by default (show them with **Columns**).
- **Account details**: the project table shows *Project name*, *Phase* and *Status*; **+ columns** adds others (remembered in your browser). Click a column title in any table to sort, click again to reverse.

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
