# Live Timeline: PowerPoint add-in (prototype)

Live Timeline redraws a clinical development timeline slide from a few inputs per scenario. Examples of inputs: study start, study length, priority review, a fixed submission date. When an input changes, every downstream date moves and the slide redraws in the same visual style as the original. That covers PCD, topline, submission, FDA review, approval and the next ACIP meeting.

- **Rules, not hand edits:** SSD + study length → PCD → topline → submission → FDA approval (8 months with priority review/PRV, 12 standard) → next ACIP meeting. Every rule can be edited in the pane.
- **Team input:** each change is logged with who made it and the source. You can copy the input table out to the team (Excel, Teams, email), paste their edited version back, and apply only the changes you accept.
- **Baseline and slippage:** save a baseline before a review. Turn on "Show slippage" to put a hollow star at the old approval date and "+3 mo" next to the new one.
- **Travels with the deck:** inputs, baseline and change log are saved inside the slide itself. The timeline is drawn as native, editable PowerPoint shapes, so people without the add-in can still open and present the slide.

## Try it without installing

Open `taskpane.html` in a browser. Outside PowerPoint it runs in preview mode on a fictional example, using the same engine and pane.

## Install (about 15 minutes, once)

An Office add-in is a small website plus a manifest file that tells PowerPoint where the website lives. The website must be served over HTTPS.

### 1. Host this folder

Any static HTTPS host works. The free route uses GitHub Pages, with no git needed:

1. On github.com, create a new **public** repository called `live-timeline`. The code holds no client data: timeline inputs live only inside your decks.
2. Click **uploading an existing file** and drag in everything from this folder, keeping the `src` and `assets` subfolders. Commit.
3. Go to **Settings › Pages**, set **Source** to *Deploy from a branch*, choose `main` and `/ (root)`, then click **Save**.
4. About a minute later the site is live at `https://YOUR-GITHUB-NAME.github.io/live-timeline/`. Open `…/taskpane.html` to check it loads the example preview.

### 2. Point the manifest at your host

Open `manifest.xml` in a text editor. Replace every `https://YOUR-HOST/live-timeline` with your address, for example `https://YOUR-GITHUB-NAME.github.io/live-timeline`, with no trailing slash. Save the file.

### 3. Load it into PowerPoint

Pick one of the following:

- **PowerPoint on the web (quickest test):** open a deck, then go to **Home › Add-ins › More Settings › Upload My Add-in** and choose `manifest.xml`. This lasts only in that browser, so re-upload if you clear the browser cache.
- **All your devices (recommended):** in the Microsoft 365 admin center, go to **Settings › Integrated apps › Add-ins › Deploy Add-in**, upload `manifest.xml` and choose **Just me**. It can take 24–72 hours to appear on the Home tab in PowerPoint for Windows, Mac and web.
- **Mac only:** copy `manifest.xml` into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` (create the `wef` folder if it is missing), then restart PowerPoint.

The add-in then shows up as **Home › Live Timeline**.

## Using it

**On a deck converted to Live Timeline** (the inputs are already saved in the slide): select the slide and click **Live Timeline**. Change an input and the slide redraws. Untick **Live** to batch several edits, then click **Update slide**.

**On a slide with a hand-drawn timeline:** the slide needs a table with a year row and a quarter row (1Q…4Q), like the filing-scenarios slide.

1. Open the pane and click **Find hand-drawn shapes**. The add-in selects the bars, markers and labels inside the table area so you can check them.
2. Click **Delete these shapes**.
3. Click **Start a live timeline for these rows** and enter the inputs.

**Tips**

- Use **−3 mo / +3 mo** on a scenario card to move a whole scenario by a quarter.
- Fill in **Source of changes** before editing (e.g. "KOL call 14"). It goes into the change log with your initials.
- Use **Team updates › Copy input table** to send the inputs out. When the edited table comes back, paste it in and click **Review changes**.
- The add-in only ever touches shapes whose names start with `LTL|`. You can see them in **Home › Arrange › Selection Pane**. Titles, the table, the rationale column, the legend and footnotes are never changed.

## Requirements

- PowerPoint for Microsoft 365 on Windows (Version 2207 or later), Mac (16.62 or later), or PowerPoint on the web.
- Reading the table's columns and rows needs Version 2508 or later (Mac 16.100+). On older builds the add-in still works on slides that already carry a saved layout, like the converted deck.
- The add-in stores everything in the presentation. It makes no network calls except loading its own files and Microsoft's Office.js library.

## Files

| File | What it is |
|---|---|
| `manifest.xml` | Tells PowerPoint where the add-in is hosted |
| `taskpane.html`, `styles.css` | The task pane |
| `src/engine.js` | Scheduling rules (dates, PRV, ACIP calendar, team-table import, baseline deltas) |
| `src/layout.js` | Turns dates into shapes in the slide's style (bars, D/S markers, review box, star, labels, Today line) |
| `src/calib.js` | Reads the time axis and scenario rows from the slide's table |
| `src/powerpoint.js` | Draws on the slide and saves the inputs in slide tags (Office.js) |
| `src/svg.js`, `src/sample.js` | Browser preview and the fictional example |
| `commands.html`, `assets/` | Ribbon support files and icons |

## Prototype status

This prototype has been checked in four ways:

1. It type-checks against Microsoft's Office.js definitions.
2. It passes tests against a strict simulation of the PowerPoint API, seeded with the real slide.
3. Its layout was rendered and compared with the original slide.
4. Its pane was driven end to end in a browser.

It has not yet been run inside a live copy of PowerPoint, so expect small visual differences: font rendering, and theme defaults on newly drawn shapes. Report anything that looks off.
