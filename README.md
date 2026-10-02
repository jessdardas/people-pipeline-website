# People Pipeline – website (IIS)

The same People Pipeline app as the Google Apps Script version, as a normal website on your Windows **IIS** server.
It is a **separate project**: nothing here touches the Apps Script project, Google Drive or clasp.

## How it works

```
data\people pipeline.xlsx  ──►  the browser reads it (SheetJS)  ──►  server\Pipeline.js links the sheets  ──►  app\ draws the page
        ▲                                                             (same file as Apps Script)             (same folder as Apps Script)
        └── tools\copy-excel.ps1, every hour (Task Scheduler)          the page reads it when it opens and when you click ↻ (never by itself)
```

| Folder / file | What it is |
|---|---|
| `index.html` | start page: puts `app\index.html` together in the browser (like Apps Script does on its server) |
| `app\` | **copy** of the Apps Script project's `app\` folder: the page, its styles and code |
| `server\Pipeline.js` | **copy** of the Apps Script project's `server\Pipeline.js`: how the Excel sheets are read and linked |
| `site\source.js` | the website's own data code: reads the Excel file from `data\` (keeps a copy in the browser for fast opening), reads `data\refresh-log.txt`, talks to `api\settings.ashx` |
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

## Fresh data

1. Open `tools\copy-excel.ps1` and set `$Source` (where your query export saves the Excel file) and `$Target`.
2. **Task Scheduler** → *Create Task* → Triggers: *Daily*, repeat every **1 hour** → Actions: *Start a program*
   `powershell.exe` with arguments `-ExecutionPolicy Bypass -File "C:\inetpub\people-pipeline\tools\copy-excel.ps1"`.

**Last updated at** (top of the page) is the last date written in `data\refresh-log.txt` (`C:\inetpub\people-pipeline\data\refresh-log.txt`). Without that file it shows when the Excel file last changed.

**The page never refreshes by itself.** The data is read when the page opens and when you click **↻** next to *Last updated at* (it then says *No changes* or *Updated*).

**Fast opening:** the browser keeps the data it read, together with the Excel file's date and size (IndexedDB, `site\source.js`). When you open the page and the file on the server has not changed, the kept copy is used: no download, no decoding (about 0.3 s instead of 1–3 s). When the file changed, it is read again and kept. Other speed-ups: the Excel reader (`site\vendor\xlsx.full.min.js`) is only loaded when the file must be decoded, decoding skips styles and formats, and the Google font no longer holds up the page (it loads in the background; without internet the normal system font is used).
If you change how the file is read (`server\Pipeline.js`), raise `cacheVersion` in `site\source.js` so every browser reads the file again once.

## Website-only features

These are only in the website's `app\` (not in the Apps Script project).

### Maps 01 and 02: two small switches (above the matrix, kept quiet on purpose)

- **Without P0** (on by default): P0 projects are left out; on map 01 the P0 column disappears. An account whose projects are all P0 leaves the map, so the totals go down. (`P0_PHASE` in `app\core\config-js.html`)
- **Real projects only** (on by default) + the small **▾** menu, where each kind of project can be ticked (left out) or unticked (shown):

  | Kind | How it is recognised | Standard |
  |---|---|---|
  | I.N. | *I.N.* in the project name (or the account name) | left out |
  | Intercompany | *intercompany* in the project name (or the account name) | left out |
  | No P2 | *pslab_nop2* of the projects sheet is TRUE | left out |
  | Excluded from pipeline | *pslab_excludingfrompipeline* of the projects sheet is TRUE | left out |
  | Internal owner | the project owner (*OwnerName*) has a word starting with ROLAND, CYBEL, ROGER, SAM, DANY DAABOUL or PSLAB (so *Cybelle*, *Sammy* and the *pslab-…* teams count too) | left out |

  Standard: all kinds ticked (left out). Unticked (shown) I.N. / intercompany / excluded / internal-owner projects go in the **other (i.n. ...)** column of map 01; shown No P2 projects stay in their phase column. An account whose projects are all left out leaves the map. Switching *Real projects only* off shows every kind. *standard* in the ▾ menu goes back to the table above. The subtitle under the map title only mentions these switches when they differ from the standard. Settings: `PROJECT_EXCLUDE`, `INTERNAL_KINDS`, `OWNER_COLS` in `app\core\config-js.html`.

### Under the maps: Account listing | Stay in touch policy (`app\touch\`)

The list of accounts under every map (after clicking a box, a row or a column title) has two views, switched at the top left of the list:

- **Account listing** – the table as before.
- **Stay in touch policy** – the same accounts, one row each (no export here, and no touchpoint columns: the touchpoints are in the pop-up and in the account details):
  - one column per week: **5 weeks before** this week, **this week** (highlighted), **10 weeks after**. Past weeks show the **subject** of each of the account's activities from the **activities** sheet, **coloured by its type** (legend above the table; hover for date · type · subject). **Tasks are never shown.** The next weeks are empty for now: later they will show the to-dos.
  - **Click an account** → the touchpoints pop-up, like a to-do list: ☑ / ☐ Site visit (last date), Pub 1, Pub 2, Pub 3 … each its own checkpoint (with the details), then *What they did* (the activities, newest first) and *Next* (the to-dos, later). *Open account details* goes to the details panel.

Where it comes from: the **activities** sheet is found by its name; its date, type and subject columns by name (`ACT_COLS`). Hidden types: `ACT_HIDE` (tasks). Colours: `ACT_COLORS` / `ACT_OTHER_COLORS`. Site visit: an account column with *site visit* in its name, or an activity whose type or subject says *site visit* (the latest date is shown). All in `app\core\config-js.html`. The account details panel shows the same touchpoints and activities.

### Account table (under the maps)

The tools above the list: on the **left** the two views, the search and **Columns**; on the **right**, away from them, the **check to export** switch with **Excel** / **PDF**. The table has no touchpoint columns (they are in the Stay in touch pop-up and the account details).

- **check to export** (right): switch it on, then tick accounts in the **#** column (the box in the # title ticks the whole list) – in either view; **Excel** / **PDF** then export only the ticked accounts. Ticks stay while the page is open (*clear* removes them).
- Stay in touch policy has only the views and the search: no export, no check boxes.
- **City** next to *Country* (the *city* column of the accounts sheet), *Last contact* next to *Last meeting*; *Phase* and *Industry* hidden by default (show them with **Columns**).

### Other

- **Search** (top): picking an account opens its details; the map stays as it is (*Show on map* in the details jumps to it).
- **Last project**: when the file has no last-project date, the latest date of the account's own projects is used: the project's *DateIn*, else its *actualclosedate* (map 02 and the details, which then say *from its projects*). `PROJECT_DATE_COLS` in `app\core\config-js.html`.
- **New DB**: with New DB on, a small date box next to *validated on or after* changes the date for your view only; nothing is saved (*reset* goes back to the admin date).
- **Account details** (a wider panel): *Website* in the overview (a link; the account column *websiteurl* / *website*); *Contact persons* as names only; the project table shows *Project name*, *Phase* and *Status* (**+ columns** adds others, remembered in your browser); *Activities* as a tidy table – Date, Type (colour), Subject, then the other filled-in columns – that scrolls sideways when it is wider than the panel. Click a column title in any table to sort, click again to reverse.

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
