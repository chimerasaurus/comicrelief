# comicrelief

An interactive Python script that fixes metadata in digital comic book files (CBZ/CBR).

Comic readers like Komga, Kavita, and YACReader rely on embedded metadata to display series names, issue numbers, publication dates, and reading order. If your files have inconsistent, missing, or wrong metadata, comics can appear out of order, grouped into phantom series, or missing covers and descriptions.

`comicrelief` scans a folder (or a single file), looks up the correct metadata from [Comic Vine](https://comicvine.gamespot.com/api/) (with [Metron](https://metron.cloud/) as a fallback), and shows you a before/after comparison before touching anything.

---

## Requirements

- Python 3.9+
- A free [Comic Vine API key](https://comicvine.gamespot.com/api/)
- For CBR support: `brew install libarchive` (macOS) or `apt install libarchive-tools` (Linux)

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

### Options

| Option | Description |
|---|---|
| `--dry-run` | Show proposed changes without writing anything |
| `--no-rename` | Fix embedded metadata only, do not rename files |
| `--no-cache` | Ignore cached API results and re-fetch from Comic Vine |
| `--api-key KEY` | Comic Vine API key (overrides env var and saved config) |
| `--cache-file FILE` | Path to JSON cache file (default: `~/.comicrelief_cache.json`) |

---

## Examples

### Preview changes without modifying any files

```bash
python3 comicrelief.py --dry-run "/Volumes/library/Fiction/Comics"
```

### Fix an entire comics library

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

### Fix metadata only — don't rename files

```bash
python3 comicrelief.py --no-rename "/Volumes/library/Fiction/Comics"
```

### Re-fetch from Comic Vine, ignoring cached results

Useful if a previous run cached a wrong match (e.g. picked the wrong issue from a multi-issue picker).

```bash
python3 comicrelief.py --no-cache "/Volumes/library/Fiction/Comics/Batman.001.cbz"
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

### 4. Confirmation UI

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

Apply changes? [y/n/q] (y):
```

**Prompt options:**
- `y` — apply the proposed metadata and rename the file
- `n` — skip this file
- `q` — quit immediately (files already processed are saved)

If no changes are detected, you get:
```
No changes detected.
  s = skip   r = re-search (pick a different match)   q = quit
```

Use `r` to bypass the cache and re-query Comic Vine — useful when the wrong issue was cached from a previous run.

### 5. Multiple issues with the same number

Some series have more than one issue with the same number (e.g. a regular edition and a variant language edition). When this happens, a picker is shown before the confirmation table:

```
Multiple issues found with this number — please choose:

  #   Title                               Cover date    ID
  1   Cadet Challenge - Klingon Edition   1998-05-31    44916
  2   Cadet Challenge                     1998-05-01    982658

Enter number [1/2] (1):
```

### 6. Writing changes

Changes are written atomically — the script builds the updated archive in a temporary file and swaps it into place, so a crash mid-write won't corrupt your file.

Files are renamed to a canonical format:
```
Series Name (Year) #001.cbz
```

Use `--no-rename` to skip renaming and only update the embedded metadata.

### 7. CBR files

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
