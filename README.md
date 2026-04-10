# comicrelief

An interactive Python script that fixes metadata in digital comic book files (CBZ/CBR).

Comic readers like Komga, Kavita, and YACReader rely on embedded metadata to display series names, issue numbers, publication dates, and reading order. If your files have inconsistent, missing, or wrong metadata, comics can appear out of order, grouped into phantom series, or missing covers and descriptions.

`comicrelief` scans a folder (or a single file), looks up the correct metadata from [Comic Vine](https://comicvine.gamespot.com/api/) (with [Metron](https://metron.cloud/) as a fallback), and shows you a before/after comparison before touching anything.

---

## Requirements

- Python 3.9+
- A free [Comic Vine API key](https://comicvine.gamespot.com/api/)
- For CBR support: `brew install libarchive` (macOS) or `apt install libarchive-tools` (Linux)
- For smart cover matching (installed automatically via requirements.txt): `Pillow` and `imagehash`

---

## Installation

```bash
git clone https://github.com/yourname/comicrelief
cd comicrelief
pip install -r requirements.txt
```

**requirements.txt**
```
requests>=2.31.0
rich>=13.7.0
rarfile>=4.1
Pillow>=10.0.0
imagehash>=4.3.1
```

---

## API Key

Get a free key at https://comicvine.gamespot.com/api/ (requires a free account).

Set it once as an environment variable:

```bash
export COMICVINE_API_KEY=a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2
```

Or pass it directly on the command line:

```bash
python3 comicrelief.py --api-key a1b2c3d4e5f6a1b2c3d4e5f6a1b2c3d4e5f6a1b2 /path/to/comics
```

Or just run the script without setting it — you'll be prompted to enter it interactively, and it will be saved to `~/.comicrelief` for future runs.

---

## Usage

```
python3 comicrelief.py [OPTIONS] <path>
```

`<path>` can be a **directory** (scanned recursively) or a **single file**.

### Modes (mutually exclusive)

| Flag | Description |
|---|---|
| _(none)_ | Interactive mode — confirm each file before applying |
| `--list` | Display a metadata summary table for all files; no changes made |
| `--convert-cbr` | Convert all CBR files to CBZ without touching metadata |
| `--auto` | Apply all changes without per-file prompts; print a change log at the end |
| `--check-pages` | Check actual image counts against stored and Comic Vine page counts |

### Options

| Option | Description |
|---|---|
| `--dry-run` | Show proposed changes without writing anything (works with all modes) |
| `--no-rename` | Fix embedded metadata only, do not rename files |
| `--no-cache` | Ignore cached API results and re-fetch from Comic Vine |
| `--smart-match` | Use cover image comparison to disambiguate series with similar names |
| `--full-metadata` | Fetch and store all available fields: Genre, Tags, LanguageISO, all characters, summary fallback from volume description |
| `--api-key KEY` | Comic Vine API key (overrides env var and saved config) |
| `--cache-file FILE` | Path to JSON cache file (default: `~/.comicrelief_cache.json`) |
| `--core-fields FIELDS` | Comma-separated fields that must all be present for a ✓ in `--list` mode. Default: `Number,Publisher,Series,Year` |
| `--fields FIELDS` | Comma-separated columns to show in `--list` mode, overriding the defaults |
| `--volume-id ID` | Force a specific Comic Vine volume ID for every file in this run |

---

## Examples

### Inspect a folder — see what needs fixing

```bash
python3 comicrelief.py --list "/Volumes/library/Fiction/Comics/Star.Trek.Starfleet.Academy.(1996)"
```

Displays a table of every file's current embedded metadata with a ✓/✗ health indicator. A ✗ means one or more core fields (Series, Number, Year, Publisher) are missing. No files are modified.

```
╭───┬──────────────────────────────────┬─────┬──────────────────────────┬───┬─────┬──────┬───────────┬──────────┬───────╮
│   │ File                             │ Fmt │ Series                   │ # │ Vol │ Year │ Publisher │ Writer   │ Pages │
├───┼──────────────────────────────────┼─────┼──────────────────────────┼───┼─────┼──────┼───────────┼──────────┼───────┤
│ ✓ │ Star_Trek-SFA.issue-001.cbz      │ CBZ │ Star Trek: Starfl…       │ 1 │ 1   │ 1996 │ Marvel    │ —        │ 30    │
│ ✗ │ Star_Trek-SFA.issue-002.cbz      │ CBZ │ —                        │ — │ —   │ —    │ —         │ —        │ —     │
│ ✗ │ Star_Trek-SFA.issue-003.cbr      │ CBR │ Star Trek SFA            │ 3 │ —   │ —    │ —         │ —        │ 28    │
╰───┴──────────────────────────────────┴─────┴──────────────────────────┴───┴─────┴──────┴───────────┴──────────┴───────╯

3 file(s)  ✓ 1 complete   ✗ 2 missing core fields (Number / Publisher / Series / Year)
```

After the per-file table, a **per-series collection summary** is printed showing how many issues are in the folder vs. the total reported by Comic Vine, with gap detection:

```
╭──────────────────────────────────┬───────────┬──────┬──────────────────────────────╮
│ Series                           │ Publisher │ Have │ Missing                      │
├──────────────────────────────────┼───────────┼──────┼──────────────────────────────┤
│ Star Trek: Starfleet Academy     │ Marvel    │ 3/19 │ #2–17, #19                   │
╰──────────────────────────────────┴───────────┴──────┴──────────────────────────────╯
```

Variant issues (e.g. an English and a Klingon edition of the same issue number) are counted once in the "Have" total.

### Customise which columns appear

Show only series, issue number, and title:

```bash
python3 comicrelief.py --list --fields "Series,Number,Title" "/Volumes/library/Fiction/Comics"
```

The default columns are: `Series`, `#` (Number), `Vol` (Volume), `Year`, `Publisher`, `Writer`, `Pages` (PageCount).

Available field names for `--fields`: `Series`, `Title`, `Number`, `Volume`, `Year`, `Month`, `Publisher`, `Imprint`, `Writer`, `Penciller`, `Inker`, `Colorist`, `CoverArtist`, `Editor`, `Genre`, `Tags`, `Characters`, `Summary`, `AgeRating`, `Count`, `PageCount`, `LanguageISO`, `StoryArc`, `Format`.

### Change what counts as "complete"

By default, a file needs Series, Number, Year, and Publisher to earn a ✓. Override this:

```bash
python3 comicrelief.py --list --core-fields "Series,Number,Title" "/Volumes/library/Fiction/Comics"
```

If a field is in `--core-fields` but not in the default columns, it is automatically added as a column so you can see the values that are being checked.

### Fix an entire comics library (interactive)

```bash
python3 comicrelief.py "/Volumes/library/Fiction/Comics"
```

### Fix a single series folder

```bash
python3 comicrelief.py "/Volumes/library/Fiction/Comics/Star.Trek.Starfleet.Academy.(1996)"
```

### Fix a single file

```bash
python3 comicrelief.py "/Volumes/library/Fiction/Comics/Star.Trek.Starfleet.Academy.(1996)/Star_Trek-Starfleet_Academy.issue-018.Marvel.1998.cbz"
```

### Preview changes without modifying any files

```bash
python3 comicrelief.py --dry-run "/Volumes/library/Fiction/Comics"
```

### Fix metadata only — don't rename files

```bash
python3 comicrelief.py --no-rename "/Volumes/library/Fiction/Comics"
```

### Re-fetch from Comic Vine, ignoring cached results

Useful if a previous run cached a wrong match.

```bash
python3 comicrelief.py --no-cache "/Volumes/library/Fiction/Comics/Batman.001.cbz"
```

### Apply all changes automatically — no prompts

```bash
python3 comicrelief.py --auto "/Volumes/library/Fiction/Comics/Star.Trek.Starfleet.Academy.(1996)"
```

You can preview what auto mode *would* do without writing anything:

```bash
python3 comicrelief.py --auto --dry-run "/Volumes/library/Fiction/Comics"
```

### Force a specific Comic Vine series

When auto-matching picks the wrong series (e.g. a foreign reprint instead of the original), find the correct series ID in its Comic Vine URL (`comicvine.gamespot.com/…/4050-**5153**/`) and pass it directly:

```bash
python3 comicrelief.py --volume-id 5153 "/Volumes/library/Fiction/Comics/Star.Trek.Deep.Space.Nine.(1993)"
```

This overrides auto-matching for every file in the run. Works with both interactive and `--auto` mode.

### Bulk-convert all CBR files to CBZ

Converts the container format only — no image re-encoding, no metadata changes.

```bash
python3 comicrelief.py --convert-cbr "/Volumes/library/Fiction/Comics"
```

Preview which files would be converted:

```bash
python3 comicrelief.py --convert-cbr --dry-run "/Volumes/library/Fiction/Comics"
```

---

## How it works

### 1. Discovery

The script recursively finds all `.cbz` and `.cbr` files under the given path.

### 2. Metadata inference

For each file it:
- Reads any existing `ComicInfo.xml` embedded in the archive
- Parses the filename to extract series name, issue number, volume, and year

Filename patterns it understands:
```
Batman.001.cbz                          → Batman #1
Star.Trek.Starfleet.Academy.(1996).001  → Star Trek Starfleet Academy (1996) #1
Batman_v1_001.cbz                       → Batman vol 1 #1
Daredevil #158.cbz                      → Daredevil #158
```

Existing metadata always takes priority over filename inference.

### 3. API lookup

Searches **Comic Vine** for the series and issue. If Comic Vine returns nothing, falls back to **Metron**.

When searching for a series, results are scored and ranked by:
- Exact name match
- Start year match
- Issue count (longer runs score higher)
- Known English publisher (DC Comics, Marvel, Image, Dark Horse, IDW, etc. score significantly higher to avoid matching foreign reprints)

API results are cached to `~/.comicrelief_cache.json` to avoid redundant requests when running over a large library. Use `--no-cache` to bypass this.

### 4. Smart cover matching

When Comic Vine returns multiple candidate series with the same or similar name (e.g. "Star Trek: Deep Space Nine" published by both Marvel and Malibu in 1993), the script extracts the cover image from the comic file and compares it against the issue cover from each candidate on Comic Vine using **perceptual hashing (pHash)**.

pHash converts an image into a 64-bit fingerprint that is robust to resizing, minor colour differences, and JPEG artefacts. The Hamming distance between two hashes indicates how similar they are (0 = identical, 64 = completely different). The candidate with the lowest distance wins.

If the best distance is above the confidence threshold (indicating no good visual match — e.g. the issue has no cover on Comic Vine, or the scan is too different from the reference), the script falls back to the score-based result.

When smart matching is used, the metadata source is shown as **Comic Vine (smart match)** in the confirmation UI.

Smart matching is **opt-in** — pass `--smart-match` to enable it. It requires `Pillow` and `imagehash` (both included in `requirements.txt`). If they are not installed, or if the cover image cannot be decoded, the script falls back to score-based matching without any error.

### 5. Confirmation UI

For each file, a two-column table is displayed showing current vs proposed metadata. Changed fields are highlighted in green, removed fields in red:

```
────────────────────────────────────────────────────────────

FILE: Batman.001.cbz
Path: /Volumes/library/Fiction/Comics/Batman.001.cbz
Metadata source: Comic Vine

╭───────────┬───────────────────────┬───────────────────────╮
│ Field     │ Current               │ Proposed              │
├───────────┼───────────────────────┼───────────────────────┤
│ Series    │ Batman                │ Batman                │
│ Number    │ 1                     │ 1                     │
│ Publisher │ (empty)               │ DC Comics             │
│ Volume    │ (empty)               │ 1                     │
│ Count     │ (empty)               │ 716                   │
│ Year      │ (empty)               │ 1940                  │
│ Month     │ (empty)               │ 4                     │
│ Summary   │ (empty)               │ The Legend of the...  │
│ PageCount │ (empty)               │ 36                    │
╰───────────┴───────────────────────┴───────────────────────╯

Apply changes? [y/n/s/i/q] (y):
```

**Prompt options:**
| Key | Action |
|---|---|
| `y` | Apply the proposed metadata and rename the file |
| `n` | Skip this file |
| `s` | Search Comic Vine for a different series name |
| `i` | Enter a Comic Vine volume ID directly |
| `q` | Quit immediately (files already processed are saved) |

If no changes are detected between current and proposed metadata, you get a different prompt:

```
No changes detected.
  s = skip   r = re-search (pick a different match)   n = new series search   i = enter ID   q = quit
```

| Key | Action |
|---|---|
| `s` | Skip this file |
| `r` | Re-query Comic Vine from scratch (bypasses cache) |
| `n` | Search for a different series name |
| `i` | Enter a Comic Vine volume ID directly |
| `q` | Quit |

### 6. Picking the right series

If the auto-scored match doesn't look right, press `s` or `n` to search for a different series name. The script queries Comic Vine and shows all results in a numbered picker:

```
Multiple volumes found — please choose:

  #   Name                                    Year   Publisher         Issues
  1   Star Trek: Deep Space Nine              1993   Marvel            32
  2   Star Trek: Deep Space Nine              1996   Marvel            15
  3   Star Trek: Deep Space Nine - Malibu     1993   Malibu Comics     28

Enter number [1/2/3] (1):
```

The chosen series sticks for all remaining files in the same run — you only need to pick once per series.

Alternatively, look up the series ID on Comic Vine (`comicvine.gamespot.com/…/4050-**ID**/`) and press `i` to enter it directly, or pass `--volume-id ID` on the command line.

### 7. Multiple issues with the same number

Some series have more than one issue with the same number (e.g. a regular edition and a variant language edition). When this happens, a picker is shown before the confirmation table:

```
Multiple issues found with this number — please choose:

  #   Title                               Cover date    ID
  1   Cadet Challenge - Klingon Edition   1998-05-31    44916
  2   Cadet Challenge                     1998-05-01    982658

Enter number [1/2] (1):
```

### 8. Writing changes

Changes are written atomically — the script builds the updated archive in a temporary file and swaps it into place, so a crash mid-write won't corrupt your file.

Files are renamed to a canonical format:
```
Series Name (Year) #001.cbz
Series Name (Year) #001 - Issue Title.cbz   ← when the issue has a title
```

Use `--no-rename` to skip renaming and only update the embedded metadata.

### 9. CBR files

CBR files (RAR archives) cannot have their contents rewritten in-place because the RAR format is proprietary. When a CBR is encountered, the script offers to convert it to CBZ first:

```
CBR file detected: Batman_001.cbr
RAR archives can't have metadata rewritten in-place. Convert to CBZ first?
Convert to CBZ? [y/n/q] (y):
```

Conversion extracts the raw image bytes and repacks them into a ZIP archive. **No image re-encoding or quality loss occurs** — the images are moved byte-for-byte, only the container format changes.

The script auto-detects available extraction tools in this order:
1. `unrar`
2. Homebrew `bsdtar` (`/opt/homebrew/opt/libarchive/bin/bsdtar`)
3. System `bsdtar`
4. `7z`

On macOS with Homebrew, `brew install libarchive` is the easiest option.

### 11. Page count checking

```bash
python3 comicrelief.py --check-pages "/Volumes/library/Fiction/Comics/Alien.(2024)"
```

Checks every CBZ/CBR in the folder and shows a table with four key values per file:

| Column | Meaning |
|---|---|
| **Actual** | Number of image files physically inside the archive |
| **Stored** | `PageCount` field in the embedded ComicInfo.xml |
| **CV** | Page count from Comic Vine (read from local cache) |
| **Delta** | Actual − CV (negative = pages possibly missing from scan) |

```
╭───┬─────────────────────────────────────┬─────┬────────┬────────┬────┬───────╮
│   │ File                                │ Fmt │ Actual │ Stored │ CV │ Delta │
├───┼─────────────────────────────────────┼─────┼────────┼────────┼────┼───────┤
│ ✓ │ Alien (2024) #001 - Bound to E….cbz │ CBZ │     23 │     23 │ 23 │     0 │
│ ✗ │ Alien (2024) #002 - Bound to E….cbz │ CBZ │     18 │     18 │ 23 │    -5 │
│ ? │ Alien (2024) #003.cbz               │ CBZ │     21 │      — │  — │     — │
╰───┴─────────────────────────────────────┴─────┴────────┴────────┴────┴───────╯

3 file(s)   ✓ 1 OK   ✗ 1 mismatch   ? 1 no reference
```

The CV column is populated from the **local cache** — no API calls are made. Run `comicrelief` on the folder first to fetch and cache metadata, then run `--check-pages` to verify your scans.

> **Note:** CV page counts include ads and letters pages. Scans that skip ads will show a small negative delta (typically −1 to −5) even when otherwise complete.

### 10. Volume inconsistency detection

The `--list` table highlights the **Vol** column in yellow when a value looks suspicious:

- The Volume equals the issue Number (a common import error where Vol was set to the issue number instead of the series volume)
- The Volume differs from the dominant value for that series across all files in the folder
- The Volume is missing while other files in the series have one

A yellow ⚠ note is printed below the table when any suspicious values are found.

---

## Metadata format

The script reads and writes **ComicInfo.xml**, the de facto standard for comic book metadata. It is placed at the root of the ZIP/CBZ archive and is supported by all major comic readers (Komga, Kavita, Ubooquity, YACReader, CDisplayEx, and more).

Fields written:

| Field | Description |
|---|---|
| `Series` | Series name |
| `Title` | Issue title |
| `Number` | Issue number |
| `Volume` | Volume number |
| `Year` / `Month` | Cover date |
| `Publisher` | Publisher name |
| `Writer` | Writer(s), comma-separated |
| `Penciller` | Penciller(s) |
| `Inker` | Inker(s) |
| `Colorist` | Colorist(s) |
| `CoverArtist` | Cover artist(s) |
| `Editor` | Editor(s) |
| `Characters` | Characters featured |
| `StoryArc` | Story arc name |
| `Summary` | Issue description |
| `Genre` | Genre |
| `AgeRating` | Age rating |
| `Count` | Total issues in series |
| `PageCount` | Number of pages (counted from archive) |
| `LanguageISO` | Language code (e.g. `en`) |

---

## Summary report

After processing, a summary is printed:

```
────────────────────────────────────────────────────────────

Summary
  Processed   42
  Updated     38
  No change    2
  Skipped      1
  Errors       1

Files with errors:
  • /Volumes/library/Fiction/Comics/corrupt_file.cbz

Files with ambiguous/missing metadata:
  • /Volumes/library/Fiction/Comics/Unknown_Comic_v2_047.cbz
```

---

## Caching

API responses are cached in `~/.comicrelief_cache.json`. This means:

- Re-running over a large library won't re-hit the API for series already looked up
- If a wrong match was cached, use `--no-cache` or press `r` at the "no changes" prompt to re-fetch

To clear the cache entirely:
```bash
rm ~/.comicrelief_cache.json
```
