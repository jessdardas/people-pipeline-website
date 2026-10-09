# People Pipeline – website (IIS)

The same People Pipeline app as the Google Apps Script version, as a normal website on your Windows **IIS** server.
It is a **separate project**: nothing here touches the Apps Script project, Google Drive or clasp.

## How it works

```
data\people pipeline.xlsx  ──►  the browser reads it (SheetJS)  ──►  server\Pipeline.js links the sheets  ──►  app\ draws the page
        ▲                                                             (same file as Apps Script)             (same folder as Apps Script)
        └── tools\copy-excel.ps1, every hour (Task Scheduler)          the page reads it when it opens (never by itself)
```

| Folder / file | What it is |
|---|---|
| `index.html` | start page: puts `app\index.html` together in the browser (like Apps Script does on its server) |
| `app\` | **copy** of the Apps Script project's `app\` folder: the page, its styles and code |
| `server\Pipeline.js` | **copy** of the Apps Script project's `server\Pipeline.js`: how the Excel sheets are read and linked |
| `site\source.js` | the website's own data code: reads the Excel file from `data\` (keeps a copy in the browser for fast opening), reads `data\refresh-log.txt`, talks to `api\settings.ashx` |
| `site\vendor\` | SheetJS, xlsx-js-style, jsPDF (+ autotable): read Excel, export Excel / PDF, without internet |
| `data\people pipeline.xlsx` | **the data**: put the Excel file here – **never in git** (`.gitignore`) |
| `api\settings.ashx` | saves the **New DB** date for everyone after the admin password (needs ASP.NET, below) |
| `App_Data\settings.json` | the saved New DB date (made by the handler, not downloadable) |
| `web.config` | IIS settings: start page, `.json` / `.xlsx` allowed, no caching, hides `App_Data` and `tools` |
| `tools\copy-excel.ps1` | copies the newest Excel file into `data\` (run it every hour) |
| `tools\update-from-app.ps1` | copies the latest `app\` + `server\Pipeline.js` from the Apps Script project |

**Which version is running?** The very small grey text at the bottom left of the page (e.g. *version 2026-10-02 · stay in touch: weeks only*) is `version` in `site\source.js`; it changes with every update. If it is older than the latest update on GitHub, the server has not been updated yet (`git pull` in the site folder, then Ctrl+F5).

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

**Last updated at** (top of the page) is the newest **successful** refresh in `data\refresh-log.txt` (`C:\inetpub\people-pipeline\data\refresh-log.txt`): only lines with a date **and the word OK** count (e.g. `2026-10-08 09:00:12 OK`); a line with **ERROR** is ignored completely, and error times are never shown. Without an OK line it shows when the Excel file last changed. (`LOG_OK` / `LOG_ERROR` in `site\source.js`.)

**The page never refreshes by itself, and there is no refresh button**: the data is read when the page opens.

**Fast opening:** the browser keeps the data it read, together with the Excel file's date and size (IndexedDB, `site\source.js`). When you open the page and the file on the server has not changed, the kept copy is used: no download, no decoding (about 0.3 s instead of 1–3 s). When the file changed, it is read again and kept. Other speed-ups: the Excel reader (`site\vendor\xlsx.full.min.js`) is only loaded when the file must be decoded, decoding skips styles and formats, and the Google font no longer holds up the page (it loads in the background; without internet the normal system font is used).
If you change how the file is read (`server\Pipeline.js`), raise `cacheVersion` in `site\source.js` so every browser reads the file again once.

## Website-only features

These are only in the website's `app\` (not in the Apps Script project).

### Navigation in 4 levels (`app\header\`)

1. **BUSINESS | NO BUSINESS** – two minimal boxes (thin border, small rounded corners); the selected one has a dark border and dark text. A divider and clear space separate them from the maps.
2. **The maps** – boxes in the same minimal look (name + number of accounts):
   - *Business*: **01 Current pipeline** (without current demand), **06 Current demand** (in the file: *current pipeline* accounts with classification order 2 – they only have P0 projects; `CURRENT_DEMAND` in `app\core\config-js.html`), **02 Past pipeline** (past clients, last project 0–18 months) and **05 Past clients** (last project **18+ months**, in ranges like Past pipeline: *1.5–2 years*, *2–3 years*, *3–4 years*, *4+ years*; rows = last meeting).
   - **Pairs together**: click *Current pipeline* and *Current demand* (or *Past pipeline* and *Past clients*) to see both **in one map** (the same layout; for the past pair, the columns of both: 0–6 months … 4+ years). Click one of the two again to keep only the other.
   - *No business*: a box per **classification** (Past demand, Met in, Met out, Contacted, Not contacted). All are shown at first; click one to see only it, click others to add them. **The map keeps its size**: a classification that is not picked stays as an empty column. (There is no *All accounts* view any more.)
3. **Service unit** and 4. **User** – the same look (no borders), the users right below the service units, starting at the same place.

On the right: the two small switches and **Filters**. **Every total of the matrix is clickable** (row total, column total and the big total) and lists those accounts below.

### The small switch (kept quiet on purpose)

(The *Without P0* switch is gone: P0-only accounts are now in their own map, 06 Current demand.)

- **Real projects only** (on by default, **the whole dashboard**) + the small **▾** menu, where each kind of project can be ticked (left out) or unticked (shown):

  | Kind | How it is recognised | Standard |
  |---|---|---|
  | I.N. | *I.N.* in the project name (or the account name) | left out |
  | Intercompany | *intercompany* in the project name (or the account name) | left out |
  | No P2 | *pslab_nop2* of the projects sheet is TRUE | left out |
  | Excluded from pipeline | *pslab_excludingfrompipeline* of the projects sheet is TRUE | left out |
  | Internal owner | the project owner (*OwnerName*) has a word starting with ROLAND, CYBEL, ROGER, SAM, DANY DAABOUL or PSLAB (so *Cybelle*, *Sammy* and the *pslab-…* teams count too) | left out |

  The number next to each kind in the ▾ menu is always the same, ticked or not: **how many accounts it concerns** in the current selection (service unit, user, New DB, Filters) – for *Internal owner* the accounts removed completely, for the other kinds the accounts that have such projects.

  **Internal users are removed everywhere, at the data level** (while *Internal owner* is ticked): a whole **account** disappears from maps, lists, totals, search, the Users row, the Filters panel, exports and details when it has **any project owned by an internal user**, or when **its own owner is an internal person** (Roland, Cybel(le), Roger, Sam(my), Dany Daaboul). The service-unit team owners (*pslab-london*, *pslab-beirut* …) are *not* internal people – they own most accounts in the file (≈ 17,500), so those accounts stay. Activities are never hidden (only accounts and projects are removed). With the current file: 451 accounts are removed (154 owned by internal people, 362 with an internal-owner project). Settings: `ACCOUNT_EXCLUDE_KINDS`, `INTERNAL_USERS` in `app\core\config-js.html` (add `'inn'` / `'ico'` there to remove the accounts of I.N. / intercompany projects too).

  The other kinds only take the **projects** out: unticked (shown) I.N. / intercompany / excluded / internal-owner projects go in the **other (i.n. ...)** column of map 01; shown No P2 projects stay in their phase column; an account whose projects are all left out leaves the map. *standard* in the ▾ menu goes back to the table above. Settings: `PROJECT_EXCLUDE`, `INTERNAL_KINDS`, `OWNER_COLS`.

### Under the maps: Account listing | Stay in touch policy | Touchpoints (`app\touch\`)

The list of accounts under every map (after clicking a box, a row / column title or a total) has two views – three on map 01. **The layout is the same in every view**: on the left the views, the search and **Sort by**; on the right the export tools (Account listing) or the colour legend (the other views); no text above the tables, and the table always starts at the same height. **One sort for all views and maps** (standard: Account name A→Z; change it with *Sort by* or a column title – it stays when you switch).

- **Account listing** – the table.
- **Stay in touch policy** – the same accounts, one row each: the **account name and the weeks only** – **5 weeks before** this week, **this week** (highlighted), **10 weeks after**. Past weeks show the **subject** of each activity (the **activities** sheet), **coloured by its type** (legend on the right; hover for date · type · subject); tasks are never shown; the next weeks are for the to-dos later.
  - Every activity type has its own colour, the same everywhere.
  - **Week filter**: click a week title → tick activity types; only the accounts that had a ticked type that week stay (*clear* removes it).
  - **Click an account** → its account details.
- **Touchpoints** (maps 01 and 06; opened projects close again when you switch view or map, and when you come back) – **#** | **Project** | **Phase** | **Accounts**, one row per unique open (real) project; the number of projects is in the list title. **Click a project** → a table opens under it: its accounts as rows, the touchpoints as columns, a round tick box per account. Buttons **P0 … P5** and **All** choose the phase (*All* = every touchpoint); every touchpoint column has the same width, its name starts at the top, and both header rows (groups, touchpoints) have a fixed height, so nothing moves from phase to phase. The group colours are soft and muted.

**Touchpoints – one system everywhere** (Touchpoints view, project pop-up, account details; `app\touch\touchpoints-js.html`): the reference (names, groups, types, phases) is `TOUCHPOINTS` / `TP_GROUPS` in `app\core\config-js.html` – 8 groups, each with its own **soft shade** (header line, ticked boxes) and **light tint** (the lit touchpoints). All tick boxes are **round** (the type – added value / deliverable – stays in the reference but is not shown). **A tick belongs to one account on one project.** For now every box starts unticked and can be ticked / unticked **on screen only: nothing is saved, a reload clears it.** Later the state will come from the database: set `TP_SOURCE` (a function project, account, touchpoint → ticked) and `TP_READONLY` in the config; the boxes then can no longer be changed by hand. The account details show the account's real projects in **the same table** (# | Project | Phase | Accounts, click to open).

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
- **Account details** (a wider panel): **every section can be closed** with the small **▾ arrow** (in the table header for tables – the header stays – or in the section title), all open by default. **Publications** (Pub 1, Pub 2 … from the contact persons). **Projects**: the account's real projects in the same table as the Touchpoints view (they start closed) (# | Project | Phase | Accounts; click a project for its accounts × touchpoints table). With *Real projects only* on, left-out projects (I.N. …) are not listed. *Website* in the overview (a link; the account column *websiteurl* / *website*); *Contact persons* as names only; the project table shows *Project name*, *Phase* and *Status* (**+ columns** adds others, remembered in your browser); *Activities* as a tidy list, newest first: the date on the left, the type (colour) and the subject, a small line with the other filled-in columns (owner, regarding, status, location …) and long texts (description) below in 2 lines (click to see all). On top: buttons to show one type only, and *Newest first / Oldest first*. Click a column title in the other tables to sort, click again to reverse.

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
