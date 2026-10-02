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
  | Excluded from pipeline | *pslab_excludingfrompipeline* of the projects sheet is TRUE | shown in **other** |
  | Internal owner | the project owner (*OwnerName*) has a word starting with ROLAND, CYBEL, ROGER, SAM, DANY DAABOUL or PSLAB (so *Cybelle*, *Sammy* and the *pslab-…* teams count too) | shown in **other** |

  Shown I.N. / intercompany / excluded / internal-owner projects go in the **other (i.n. ...)** column of map 01; shown No P2 projects stay in their phase column. An account whose projects are all left out leaves the map. Switching *Real projects only* off shows every kind. *standard* in the ▾ menu goes back to the table above. The subtitle under the map title only mentions these switches when they differ from the standard. Settings: `PROJECT_EXCLUDE`, `INTERNAL_KINDS`, `OWNER_COLS` in `app\core\config-js.html`.

### Touchpoints (tab 05, `app\touch\`)

For the accounts of the current selection (service units, owners, New DB, Filters):

1. **Activities by week**: one column per week, **5 weeks before** this week, **this week** (highlighted) and **10 weeks after** (`WEEKS_BEFORE` / `WEEKS_AFTER`). The past weeks show the account's activities from the **activities** sheet of the Excel file (hover for date, type and subject). The next weeks are empty for now: later they will show the to-dos (what should happen next). Only accounts with activities in these weeks are listed; click a name for its details.
2. **Touchpoints**: read-only tick boxes for **Site visit** and **Publication**, with the publication details. Only accounts with a touchpoint are listed (tick *also accounts without touchpoints* for all).
   - *Site visit*: an account column with *site visit* in its name that is filled in, or an activity whose type or subject says *site visit*.
   - *Publication*: the contact person columns `pslab_publication1`, `pslab_publication2`, … A text is shown as the publication; a plain *yes* shows the column name (*Publication 2*).

The **activities** sheet is found by its name (*activities*); its date, type and subject columns are found by name (`ACT_COLS` in `app\core\config-js.html`). The account details panel also shows the touchpoints and the account's activities.

### Account table (under the maps)

- **check to export**: tick it, then tick accounts in the **#** column (the box in the # title ticks the whole table); **Excel** / **PDF** then export only the ticked accounts. Ticks stay while the page is open (*clear* removes them).
- **touchpoints**: adds the *Site visit* and *Publications* columns (off by default).
- **City** next to *Country* (the *city* column of the accounts sheet), *Last contact* next to *Last meeting*; *Phase* and *Industry* hidden by default (show them with **Columns**).

### Other

- **Search** (top): picking an account opens its details; the map stays as it is (*Show on map* in the details jumps to it).
- **Last project**: when the file has no last-project date, the latest date of the account's own projects is used: the project's *DateIn*, else its *actualclosedate* (map 02 and the details, which then say *from its projects*). `PROJECT_DATE_COLS` in `app\core\config-js.html`.
- **New DB**: with New DB on, a small date box next to *validated on or after* changes the date for your view only; nothing is saved (*reset* goes back to the admin date).
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
