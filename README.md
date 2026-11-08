# Study Tracker – Word add-in

A panel inside Word that audits your study sessions on files you highlight as you go.

- **Active time**: only counts time around your highlights (and any typing). After each highlight you're counted as active until your next highlight, up to a limit you choose (default 2 min). Read for 1 minute between highlights → the minute counts. Leave for lunch → only the first 2 minutes after your last highlight count.
- **Progress**: position of your furthest highlight ÷ total words in the file.
- **Time left**: your pace (progress gained ÷ active time) applied to what's left.
- **Sessions**: a log of each sitting (date, active minutes, % gained). A pause longer than 30 min starts a new session. Copy it all as CSV for a spreadsheet.
- **Jump to my last highlight** puts the cursor back where you stopped.

Stats are saved inside the .docx itself (save the file to keep them), so each file has its own history.

---

## 1. Host the add-in (one time)

Word add-ins are small web pages, so the `docs` folder must be reachable over https.

### Option A – GitHub Pages (free, recommended: works every day on Windows, Mac and Word on the web, no server to keep running)

1. Create a free account at github.com, then **New repository** → name it `study-tracker` → Public → Create.
2. Click **uploading an existing file** and drag in everything from this folder (including the `docs` folder). Commit.
3. Repository **Settings → Pages** → Source: **Deploy from a branch** → Branch: `main`, folder: `/docs` → Save. After a minute your add-in lives at `https://YOUR-USERNAME.github.io/study-tracker/`.
4. Open `manifest.template.xml` in Notepad/TextEdit, use Replace All to change `{{BASE}}` into `https://YOUR-USERNAME.github.io/study-tracker` (no slash at the end) and save it as `manifest.xml`.
   (With Node.js installed you can instead run `node set-url.js https://YOUR-USERNAME.github.io/study-tracker`.)

Only the add-in's code is public. Your documents and study data never leave your computer/OneDrive.

### Option B – run it on your own computer (needs Node.js; the window must stay open while you study)

```
npm install
npm start          # serves the add-in at https://localhost:3000 and trusts a local certificate (first run asks permission)
```

The included `manifest.xml` already points to `https://localhost:3000`.

## 2. Add it to Word (one time)

- **Word on the web**: Home → Add-ins → More Add-ins → My Add-ins → **Upload My Add-in** → choose `manifest.xml`.
- **Mac**: copy `manifest.xml` into `~/Library/Containers/com.microsoft.Word/Data/Documents/wef` (create the `wef` folder if missing; in Finder use Go → Go to Folder). Restart Word.
- **Windows**:
  1. Put `manifest.xml` in a folder, e.g. `C:\AddinManifests`. Right-click the folder → Properties → Sharing → Share, and copy the network path shown (like `\\YOUR-PC\AddinManifests`).
  2. In Word: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs. Paste the network path in Catalog Url → Add catalog → tick **Show in Menu** → OK. Restart Word.
  3. Home → Add-ins → More Add-ins → **Shared Folder** tab → Study Tracker → Add.
- **Shortcut with Node.js** (Windows or Mac): `npm run sideload`.

## 3. Use it

1. Open a file you're studying, then **Home → Study Tracker** (in the "Study" group).
2. In Settings tick **Open this panel automatically when I open this file**, so you never forget to start it.
3. Highlight as you study. The clock starts with your first highlight.
4. Save the file when you're done so the stats are stored.

### Settings
- **Count me as active for up to**: how long after a highlight still counts as studying (30 s – 10 min). Pick roughly the longest you'd read without highlighting.
- **Highlights that count toward progress**: if the file came with highlights (e.g. from a professor) in another colour, pick your colour so they don't inflate your progress.
- **Pause** (top-right button) stops the clock without closing the panel.

## Good to know
- Timing only happens while the panel is open. Highlights made with the panel closed still move your progress, but add no time.
- Highlights are noticed within a few seconds (Word doesn't announce formatting changes, so the panel checks the document while you're working; a quick check every minute otherwise).
- Progress is word-based: pictures or pasted slides without text don't count toward it.
- Only Word's highlighter counts, not text shading or font colour.
- The time-left estimate appears after ~3 minutes of active study and 0.5% of progress, and gets steadier the more you study. If you change the counted colour, the pace restarts from that point.
- Needs Word 2016 or newer / Microsoft 365, or Word on the web.

## Files
- `docs/` – the add-in (task pane page, logic, icons)
- `manifest.xml` – tells Word where the add-in is (points to localhost); `manifest.template.xml` – the same with a `{{BASE}}` placeholder
- `server.js`, `package.json` – local hosting (Option B)
- `test/` – logic tests (`npm test`) and a browser test against a simulated Word document
