#!/usr/bin/env python3
"""
comicrelief.py — Comic book metadata fixer.

Scans a directory of CBZ/CBR files, fetches correct metadata from Comic Vine
(with Metron as fallback), and interactively applies fixes with a two-panel
before/after confirmation UI.

Usage:
    python comicrelief.py [OPTIONS] <directory>

Options:
    --dry-run       Show proposed changes without writing anything
    --no-rename     Do not rename files, only fix embedded metadata
    --api-key KEY   Comic Vine API key (overrides env/config)
    --cache-file F  Path to JSON cache file for API results (default: ~/.comicrelief_cache.json)
"""

import argparse
import json
import warnings
warnings.filterwarnings("ignore")  # suppress urllib3/ssl noise on macOS system Python
from typing import Optional, Tuple, List, Dict
import os
import re
import shutil
import sys
import tempfile
import time
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom

import requests
from rich.columns import Columns
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt
from rich.table import Table
from rich.text import Text
from rich import box

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

CONFIG_PATH = Path.home() / ".comicrelief"
DEFAULT_CACHE_PATH = Path.home() / ".comicrelief_cache.json"
COMICVINE_SEARCH_URL = "https://comicvine.gamespot.com/api/search/"
COMICVINE_VOLUMES_URL = "https://comicvine.gamespot.com/api/volumes/"
COMICVINE_ISSUES_URL = "https://comicvine.gamespot.com/api/issues/"
METRON_BASE_URL = "https://metron.cloud/api/v1"

# Fields we care about, in display order
METADATA_FIELDS = [
    "Series",
    "Title",
    "Number",
    "Volume",
    "Year",
    "Month",
    "Publisher",
    "Imprint",
    "Writer",
    "Penciller",
    "Inker",
    "Colorist",
    "CoverArtist",
    "Editor",
    "Genre",
    "Tags",
    "Characters",
    "Summary",
    "AgeRating",
    "Count",
    "PageCount",
    "LanguageISO",
    "StoryArc",
    "Format",
]

# Fields we should NOT carry over from existing metadata when the API returns a fresh result.
# If the API doesn't have them, blank is more honest than keeping a potentially wrong value.
FIELDS_NO_PRESERVE = {"Summary", "Title"}

console = Console()


# ---------------------------------------------------------------------------
# Config / API key management
# ---------------------------------------------------------------------------

def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text())
        except Exception:
            return {}
    return {}


def save_config(cfg: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2))


def get_api_key(args_key: Optional[str]) -> Optional[str]:
    if args_key:
        return args_key
    env_key = os.environ.get("COMICVINE_API_KEY")
    if env_key:
        return env_key
    cfg = load_config()
    if cfg.get("comicvine_api_key"):
        return cfg["comicvine_api_key"]
    return None


def prompt_for_api_key() -> str:
    console.print("\n[yellow]No Comic Vine API key found.[/yellow]")
    console.print("Get a free key at: [link]https://comicvine.gamespot.com/api/[/link]")
    key = Prompt.ask("Enter your Comic Vine API key (or press Enter to skip and use Metron only)")
    if key:
        cfg = load_config()
        cfg["comicvine_api_key"] = key
        save_config(cfg)
        console.print("[green]API key saved to ~/.comicrelief[/green]")
    return key


# ---------------------------------------------------------------------------
# Disk cache for API responses
# ---------------------------------------------------------------------------

class DiskCache:
    def __init__(self, path: Path):
        self.path = path
        self._data: dict = {}
        self._load()

    def _load(self) -> None:
        if self.path.exists():
            try:
                self._data = json.loads(self.path.read_text())
            except Exception:
                self._data = {}

    def get(self, key: str):
        return self._data.get(key)

    def set(self, key: str, value) -> None:
        self._data[key] = value
        try:
            self.path.write_text(json.dumps(self._data, indent=2))
        except Exception:
            pass  # non-fatal


# ---------------------------------------------------------------------------
# ComicInfo.xml handling
# ---------------------------------------------------------------------------

COMICINFO_NS = ""  # No namespace in ComicInfo.xml v2


def parse_comicinfo(xml_str: str) -> dict:
    """Parse ComicInfo.xml content into a flat dict."""
    result = {}
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError:
        return result
    for field in METADATA_FIELDS:
        el = root.find(field)
        if el is not None and el.text:
            result[field] = el.text.strip()
    return result


def build_comicinfo(metadata: dict) -> str:
    """Build a ComicInfo.xml string from a flat metadata dict."""
    root = ET.Element("ComicInfo")
    root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
    for field in METADATA_FIELDS:
        val = metadata.get(field)
        if val:
            el = ET.SubElement(root, field)
            el.text = str(val)
    raw = ET.tostring(root, encoding="unicode")
    # Pretty-print
    try:
        pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
        return pretty.decode("utf-8")
    except Exception:
        return raw


# ---------------------------------------------------------------------------
# Archive handling
# ---------------------------------------------------------------------------

def read_cbz_metadata(path: Path) -> dict:
    """Read ComicInfo.xml from a CBZ (zip) file."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            names_lower = {n.lower(): n for n in zf.namelist()}
            key = names_lower.get("comicinfo.xml")
            if key:
                return parse_comicinfo(zf.read(key).decode("utf-8", errors="replace"))
    except Exception as e:
        console.print(f"[red]Error reading {path.name}: {e}[/red]")
    return {}


def read_cbr_metadata(path: Path) -> dict:
    """Read ComicInfo.xml from a CBR (rar) file."""
    try:
        import rarfile
        with rarfile.RarFile(path, "r") as rf:
            names_lower = {n.lower(): n for n in rf.namelist()}
            key = names_lower.get("comicinfo.xml")
            if key:
                return parse_comicinfo(rf.read(key).decode("utf-8", errors="replace"))
    except ImportError:
        console.print("[yellow]rarfile not installed — cannot read CBR metadata.[/yellow]")
    except Exception as e:
        console.print(f"[red]Error reading {path.name}: {e}[/red]")
    return {}


def read_metadata(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext == ".cbz":
        return read_cbz_metadata(path)
    elif ext == ".cbr":
        return read_cbr_metadata(path)
    return {}


def get_page_count(path: Path) -> int:
    """Count image pages in the archive."""
    image_exts = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff"}
    try:
        if path.suffix.lower() == ".cbz":
            with zipfile.ZipFile(path, "r") as zf:
                return sum(
                    1 for n in zf.namelist()
                    if Path(n).suffix.lower() in image_exts
                )
        elif path.suffix.lower() == ".cbr":
            import rarfile
            with rarfile.RarFile(path, "r") as rf:
                return sum(
                    1 for n in rf.namelist()
                    if Path(n).suffix.lower() in image_exts
                )
    except Exception:
        pass
    return 0


def write_cbz_metadata(path: Path, metadata: dict, dry_run: bool = False) -> bool:
    """Atomically rewrite a CBZ file with updated ComicInfo.xml."""
    if dry_run:
        return True
    xml_content = build_comicinfo(metadata).encode("utf-8")
    tmp_path = path.with_suffix(".cbz.tmp")
    try:
        with zipfile.ZipFile(path, "r") as src_zf:
            with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as dst_zf:
                # Copy all entries except existing ComicInfo.xml
                for item in src_zf.infolist():
                    if item.filename.lower() == "comicinfo.xml":
                        continue
                    dst_zf.writestr(item, src_zf.read(item.filename))
                # Write new ComicInfo.xml
                dst_zf.writestr("ComicInfo.xml", xml_content)
        shutil.move(str(tmp_path), str(path))
        return True
    except Exception as e:
        console.print(f"[red]Failed to write {path.name}: {e}[/red]")
        if tmp_path.exists():
            tmp_path.unlink()
        return False


def _find_rar_tool() -> Optional[str]:
    """Return the path to an available RAR extraction tool, or None."""
    # Prefer unrar, then Homebrew's bsdtar (has RAR support), then system bsdtar, then 7z
    candidates = [
        "unrar",
        "/opt/homebrew/opt/libarchive/bin/bsdtar",  # Homebrew libarchive (macOS ARM)
        "/usr/local/opt/libarchive/bin/bsdtar",     # Homebrew libarchive (macOS Intel)
        "bsdtar",
        "7z",
    ]
    for tool in candidates:
        if shutil.which(tool):
            return tool
    return None


def convert_cbr_to_cbz(cbr_path: Path, dry_run: bool = False) -> Optional[Path]:
    """Convert a CBR to CBZ. Returns the new CBZ path, or None on failure."""
    cbz_path = cbr_path.with_suffix(".cbz")
    if dry_run:
        return cbz_path

    rar_tool = _find_rar_tool()
    if not rar_tool:
        console.print(
            "[red]No RAR extraction tool found.[/red] "
            "Install one with: [bold]brew install unar[/bold] (or brew install rar)"
        )
        return None

    # Configure rarfile to use whatever tool is available
    try:
        import rarfile
        if rar_tool == "bsdtar":
            rarfile.BSDTAR_TOOL = "bsdtar"
            rarfile.ALT_TOOL = "bsdtar"
            rarfile.CURRENT_SETUP.tool = "bsdtar"
        elif rar_tool == "7z":
            rarfile.SEVENZIP_TOOL = "7z"
    except Exception:
        pass

    # Use a temp extract dir to avoid partial writes
    import tempfile
    tmp_dir = Path(tempfile.mkdtemp())
    try:
        import subprocess

        # Extract the CBR to a temp directory using the available tool
        if rar_tool == "unrar":
            cmd = ["unrar", "x", "-o+", str(cbr_path), str(tmp_dir) + "/"]
        elif "bsdtar" in rar_tool:
            cmd = [rar_tool, "-xf", str(cbr_path), "-C", str(tmp_dir)]
        elif rar_tool == "7z":
            cmd = ["7z", "x", str(cbr_path), f"-o{tmp_dir}", "-y"]
        else:
            raise RuntimeError(f"Unknown tool: {rar_tool}")

        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip())

        # Pack extracted files into a CBZ
        image_files = sorted(
            f for f in tmp_dir.rglob("*")
            if f.is_file() and f.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff"}
        )
        other_files = sorted(
            f for f in tmp_dir.rglob("*")
            if f.is_file() and f.suffix.lower() not in {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tiff"}
               and f.name.lower() != "comicinfo.xml"
        )

        with zipfile.ZipFile(cbz_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for f in image_files + other_files:
                zf.write(f, f.relative_to(tmp_dir))

        console.print(f"[green]Converted to CBZ:[/green] {cbz_path.name}")
        return cbz_path

    except Exception as e:
        console.print(f"[red]CBR conversion failed: {e}[/red]")
        if cbz_path.exists():
            cbz_path.unlink()
        return None
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Filename parsing
# ---------------------------------------------------------------------------

def slugify_series(name: str) -> str:
    """Normalize a series name for comparison/search queries."""
    # Replace dots/underscores/hyphens with spaces
    name = re.sub(r"[._\-]+", " ", name)
    # Remove parenthesised years
    name = re.sub(r"\(\d{4}\)", "", name)
    # Remove excess whitespace
    name = re.sub(r"\s+", " ", name).strip()
    return name


def parse_filename(path: Path) -> dict:
    """
    Extract series, issue number, volume, and year from a comic filename.
    Returns a dict with whatever could be inferred (may be partial).
    """
    stem = path.stem
    result: dict = {}

    # --- Year: look for (YYYY) or _YYYY_ patterns ---
    year_match = re.search(r"\((\d{4})\)", stem)
    if not year_match:
        year_match = re.search(r"[\._\- ](\d{4})[\._\- ]", stem)
    if year_match:
        result["Year"] = year_match.group(1)

    # --- Issue number: various patterns ---
    # #NNN or #NN or #N (possibly with decimal like #1.5)
    issue_match = re.search(r"#(\d+(?:\.\d+)?)", stem)
    if not issue_match:
        # _NNN or -NNN or .NNN at end (zero-padded) — must be trailing
        issue_match = re.search(r"[\._\- ](\d{1,4})(?:[\._\- ]|$)", stem)
    if not issue_match:
        # v\d+_NNN or v\d+.NNN
        issue_match = re.search(r"v\d+[\._\- ]*(\d{1,4})(?:[\._\- ]|$)", stem)
    if issue_match:
        result["Number"] = issue_match.group(1).lstrip("0") or "0"

    # --- Volume ---
    vol_match = re.search(r"[Vv]ol(?:ume)?\.?\s*(\d+)", stem)
    if not vol_match:
        vol_match = re.search(r"[Vv](\d+)[_\- ]", stem)
    if vol_match:
        result["Volume"] = vol_match.group(1)

    # --- Series name: strip issue/volume/year tokens ---
    series = stem
    # Remove year in parens
    series = re.sub(r"\(\d{4}\)", "", series)
    # Remove #NNN
    series = re.sub(r"#\d+(?:\.\d+)?", "", series)
    # Remove Vol/v patterns
    series = re.sub(r"[Vv]ol(?:ume)?\.?\s*\d+", "", series)
    series = re.sub(r"[Vv]\d+[_\- ]", " ", series)
    # Remove trailing issue numbers: _001 or -001 or space 001
    series = re.sub(r"[\._\- ]\d{1,4}$", "", series)
    # Replace dots/underscores with spaces
    series = re.sub(r"[._]+", " ", series)
    series = re.sub(r"\s+", " ", series).strip()
    if series:
        result["Series"] = series

    return result


# ---------------------------------------------------------------------------
# Comic Vine API
# ---------------------------------------------------------------------------

COMICVINE_RATE_LIMIT = 1.0  # seconds between requests
_last_cv_request = 0.0


def _cv_get(url: str, params: dict, api_key: str) -> Optional[dict]:
    global _last_cv_request
    elapsed = time.time() - _last_cv_request
    if elapsed < COMICVINE_RATE_LIMIT:
        time.sleep(COMICVINE_RATE_LIMIT - elapsed)
    params = {**params, "api_key": api_key, "format": "json"}
    headers = {"User-Agent": "comicrelief/1.0 (comic metadata fixer)"}
    try:
        r = requests.get(url, params=params, headers=headers, timeout=15)
        _last_cv_request = time.time()
        if r.status_code == 200:
            return r.json()
        console.print(f"[yellow]Comic Vine returned HTTP {r.status_code}[/yellow]")
    except requests.RequestException as e:
        console.print(f"[yellow]Comic Vine request failed: {e}[/yellow]")
    return None


def search_comicvine_volume(series_name: str, year: Optional[str], api_key: str, cache: DiskCache) -> Optional[dict]:
    """Search for a volume (series) on Comic Vine. Returns the best matching volume dict."""
    cache_key = f"cv_vol:{series_name.lower()}:{year or ''}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    data = _cv_get(
        COMICVINE_SEARCH_URL,
        {
            "query": series_name,
            "resources": "volume",
            "field_list": "id,name,start_year,publisher,count_of_issues,description",
            "limit": 10,
        },
        api_key,
    )
    if not data or data.get("status_code") != 1:
        cache.set(cache_key, None)
        return None

    results = data.get("results", [])
    if not results:
        cache.set(cache_key, None)
        return None

    # Score candidates: exact name match + year match + issue count + English publisher
    ENGLISH_PUBLISHERS = {
        "dc comics", "marvel", "image comics", "dark horse comics", "idw publishing",
        "boom! studios", "dynamite entertainment", "vertigo", "wildstorm", "valiant",
        "archie comics", "oni press", "titan comics", "aftershock", "antarctic press",
    }

    def score(v):
        s = 0
        vname = v.get("name", "").lower()
        if vname == series_name.lower():
            s += 10
        elif series_name.lower() in vname:
            s += 5
        if year and v.get("start_year") == year:
            s += 8
        # Prefer volumes with more issues (likely the main run)
        count = v.get("count_of_issues") or 0
        if count > 100:
            s += 4
        elif count > 20:
            s += 2
        elif count > 5:
            s += 1
        # Strongly prefer known English publishers
        pub = v.get("publisher") or {}
        pub_name = (pub.get("name", "") if isinstance(pub, dict) else str(pub)).lower()
        if pub_name in ENGLISH_PUBLISHERS:
            s += 12
        return s

    best = max(results, key=score)
    cache.set(cache_key, best)
    return best


def fetch_comicvine_issue(volume_id: int, issue_number: str, api_key: str, cache: DiskCache, skip_cache: bool = False) -> Optional[dict]:
    """Fetch a specific issue from a Comic Vine volume."""
    cache_key = f"cv_issue:{volume_id}:{issue_number}"
    if not skip_cache:
        cached = cache.get(cache_key)
        if cached is not None:
            return cached

    # Normalize issue number for CV (strip leading zeros)
    normalized = str(int(float(issue_number))) if issue_number.replace(".", "").isdigit() else issue_number

    data = _cv_get(
        COMICVINE_ISSUES_URL,
        {
            "filter": f"volume:{volume_id},issue_number:{normalized}",
            "field_list": (
                "id,name,issue_number,cover_date,description,person_credits,"
                "character_credits,story_arc_credits,volume"
            ),
            "limit": 5,
        },
        api_key,
    )
    if not data or data.get("status_code") != 1:
        cache.set(cache_key, None)
        return None

    results = data.get("results", [])
    if not results:
        cache.set(cache_key, None)
        return None

    if len(results) == 1:
        issue = results[0]
        cache.set(cache_key, issue)
        return issue

    # Multiple issues share this number — ask the user to pick one
    issue = _pick_issue(results)
    cache.set(cache_key, issue)
    return issue


def _pick_issue(results: list) -> dict:
    """Show a numbered list of issues and let the user choose one."""
    console.print(f"\n[yellow]Multiple issues found with this number — please choose:[/yellow]")
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("#", width=4)
    table.add_column("Title")
    table.add_column("Cover date", width=12)
    table.add_column("ID", width=10)

    for i, issue in enumerate(results, 1):
        table.add_row(
            str(i),
            issue.get("name") or "(untitled)",
            issue.get("cover_date") or "?",
            str(issue.get("id", "")),
        )
    console.print(table)

    choices = [str(i) for i in range(1, len(results) + 1)]
    choice = Prompt.ask("Enter number", choices=choices, default="1")
    return results[int(choice) - 1]


def extract_cv_metadata(volume: dict, issue: Optional[dict]) -> dict:
    """Convert Comic Vine volume + issue dicts into our flat metadata format."""
    meta = {}

    # Series from volume
    meta["Series"] = volume.get("name", "")

    # Publisher
    pub = volume.get("publisher")
    if pub:
        meta["Publisher"] = pub.get("name", "") if isinstance(pub, dict) else str(pub)

    # Start year as volume year
    start_year = volume.get("start_year")
    if start_year:
        meta["Volume"] = "1"  # CV doesn't distinguish volume numbers well

    # Total issues
    count = volume.get("count_of_issues")
    if count:
        meta["Count"] = str(count)

    if issue:
        # Issue number
        num = issue.get("issue_number")
        if num:
            meta["Number"] = str(num).lstrip("0") or "0"

        # Issue title (name in Comic Vine)
        title = issue.get("name", "").strip()
        if title:
            meta["Title"] = title

        # Cover date → Year + Month
        cover_date = issue.get("cover_date", "")
        if cover_date:
            parts = cover_date.split("-")
            if len(parts) >= 1 and parts[0]:
                meta["Year"] = parts[0]
            if len(parts) >= 2 and parts[1]:
                meta["Month"] = parts[1].lstrip("0") or "1"

        # Title
        title = issue.get("name")
        # We don't put title into Series, but we could use it for a "Title" field
        # ComicInfo.xml has no "Title" top-level, it's inside the Series/Number combo.
        # Some readers use it; let's skip for now to keep things clean.

        # Writers and artists from person_credits
        writers, pencillers, inkers, colorists, letterers, cover_artists, editors = [], [], [], [], [], [], []
        for person in issue.get("person_credits", []):
            role = person.get("role", "").lower()
            name = person.get("name", "")
            if "writer" in role:
                writers.append(name)
            if "pencil" in role:
                pencillers.append(name)
            if "ink" in role:
                inkers.append(name)
            if "color" in role or "colour" in role:
                colorists.append(name)
            if "letter" in role:
                letterers.append(name)
            if "cover" in role:
                cover_artists.append(name)
            if "edit" in role:
                editors.append(name)

        if writers:
            meta["Writer"] = ", ".join(writers)
        if pencillers:
            meta["Penciller"] = ", ".join(pencillers)
        if inkers:
            meta["Inker"] = ", ".join(inkers)
        if colorists:
            meta["Colorist"] = ", ".join(colorists)
        if cover_artists:
            meta["CoverArtist"] = ", ".join(cover_artists)
        if editors:
            meta["Editor"] = ", ".join(editors)

        # Characters
        chars = [c.get("name", "") for c in issue.get("character_credits", [])]
        if chars:
            meta["Characters"] = ", ".join(chars[:20])  # cap at 20

        # Story arcs
        arcs = [a.get("name", "") for a in issue.get("story_arc_credits", [])]
        if arcs:
            meta["StoryArc"] = ", ".join(arcs)

        # Summary — strip HTML tags from description, normalize whitespace
        desc = issue.get("description", "") or ""
        desc = re.sub(r"<br\s*/?>", "\n", desc, flags=re.IGNORECASE)  # preserve line breaks
        desc = re.sub(r"<[^>]+>", " ", desc)   # replace other tags with space
        desc = re.sub(r"[ \t]+", " ", desc)     # collapse horizontal whitespace
        desc = re.sub(r"\n[ \t]+", "\n", desc)  # trim leading space on each line
        desc = re.sub(r"\n{3,}", "\n\n", desc)  # max two consecutive newlines
        desc = desc.strip()
        if desc:
            meta["Summary"] = desc[:2000]  # cap length

    return {k: v for k, v in meta.items() if v}


# ---------------------------------------------------------------------------
# Metron API (fallback, no key needed)
# ---------------------------------------------------------------------------

METRON_RATE_LIMIT = 1.0
_last_metron_request = 0.0


def _metron_get(endpoint: str, params: dict) -> Optional[dict]:
    global _last_metron_request
    elapsed = time.time() - _last_metron_request
    if elapsed < METRON_RATE_LIMIT:
        time.sleep(METRON_RATE_LIMIT - elapsed)
    url = f"{METRON_BASE_URL}/{endpoint}"
    try:
        r = requests.get(url, params=params, timeout=15,
                         headers={"User-Agent": "comicrelief/1.0"})
        _last_metron_request = time.time()
        if r.status_code == 200:
            return r.json()
    except requests.RequestException as e:
        console.print(f"[yellow]Metron request failed: {e}[/yellow]")
    return None


def search_metron(series_name: str, issue_number: Optional[str], year: Optional[str], cache: DiskCache) -> Optional[dict]:
    """Search Metron for a series+issue and return flat metadata dict."""
    cache_key = f"metron:{series_name.lower()}:{issue_number or ''}:{year or ''}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    # Search for the series
    data = _metron_get("series/", {"name": series_name, "page_size": 10})
    if not data or not data.get("results"):
        cache.set(cache_key, None)
        return None

    results = data["results"]

    def score(s):
        sc = 0
        if s.get("name", "").lower() == series_name.lower():
            sc += 10
        if year and s.get("year_began") == int(year):
            sc += 5
        return sc

    best_series = max(results, key=score)
    series_id = best_series.get("id")
    if not series_id:
        cache.set(cache_key, None)
        return None

    meta = {
        "Series": best_series.get("name", ""),
        "Publisher": best_series.get("publisher", {}).get("name", "") if best_series.get("publisher") else "",
    }
    if best_series.get("year_began"):
        meta["Year"] = str(best_series["year_began"])

    # Fetch the specific issue
    if issue_number:
        issue_data = _metron_get("issue/", {"series_id": series_id, "number": issue_number, "page_size": 5})
        if issue_data and issue_data.get("results"):
            issue = issue_data["results"][0]
            meta["Number"] = str(issue.get("number", issue_number))
            cover = issue.get("cover_date", "")
            if cover:
                parts = cover.split("-")
                if parts[0]:
                    meta["Year"] = parts[0]
                if len(parts) >= 2 and parts[1]:
                    meta["Month"] = parts[1].lstrip("0") or "1"
            # Credits
            writers, pencillers = [], []
            for credit in issue.get("credits", []):
                role = credit.get("role", [])
                roles = [r.get("name", "").lower() for r in role] if isinstance(role, list) else [str(role).lower()]
                name = credit.get("creator", {}).get("name", "") if isinstance(credit.get("creator"), dict) else str(credit.get("creator", ""))
                if any("writer" in r for r in roles):
                    writers.append(name)
                if any("pencil" in r for r in roles):
                    pencillers.append(name)
            if writers:
                meta["Writer"] = ", ".join(writers)
            if pencillers:
                meta["Penciller"] = ", ".join(pencillers)

    result = {k: v for k, v in meta.items() if v}
    cache.set(cache_key, result)
    return result if result else None


# ---------------------------------------------------------------------------
# Metadata lookup orchestration
# ---------------------------------------------------------------------------

def fetch_metadata(inferred: dict, api_key: Optional[str], cache: DiskCache, skip_cache: bool = False) -> Tuple[Optional[dict], str]:
    """
    Fetch metadata from Comic Vine (primary) or Metron (fallback).
    Returns (metadata_dict, source_label) or (None, "not found").
    """
    series = inferred.get("Series", "")
    if not series:
        return None, "no series name"

    search_name = slugify_series(series)
    issue_num = inferred.get("Number")
    year = inferred.get("Year")

    # --- Comic Vine ---
    if api_key:
        volume = search_comicvine_volume(search_name, year, api_key, cache)
        if volume:
            vol_id = volume.get("id")
            issue = None
            if vol_id and issue_num:
                issue = fetch_comicvine_issue(vol_id, issue_num, api_key, cache, skip_cache=skip_cache)
            meta = extract_cv_metadata(volume, issue)
            return meta, "Comic Vine"

    # --- Metron fallback ---
    meta = search_metron(search_name, issue_num, year, cache)
    if meta:
        return meta, "Metron"

    return None, "not found"


# ---------------------------------------------------------------------------
# Rich UI — two-panel confirmation
# ---------------------------------------------------------------------------

def _format_value(val: Optional[str], max_len: int = 60) -> str:
    if not val:
        return "[dim](empty)[/dim]"
    val = str(val)
    if len(val) > max_len:
        val = val[:max_len - 3] + "..."
    return val


def _field_row(field: str, old: Optional[str], new: Optional[str]) -> Tuple[str, str, str]:
    """Return (field_label, old_display, new_display) with colour markup."""
    old_val = old or ""
    new_val = new or ""
    label = field

    if old_val == new_val:
        old_str = _format_value(old_val)
        new_str = _format_value(new_val)
    elif not old_val and new_val:
        old_str = "[dim](empty)[/dim]"
        new_str = f"[green]{_format_value(new_val)}[/green]"
    elif old_val and not new_val:
        old_str = f"[red]{_format_value(old_val)}[/red]"
        new_str = "[dim](removing)[/dim]"
    else:
        old_str = f"[yellow]{_format_value(old_val)}[/yellow]"
        new_str = f"[green]{_format_value(new_val)}[/green]"

    return label, old_str, new_str


def show_confirmation_ui(
    file_path: Path,
    current: dict,
    proposed: dict,
    source: str,
    dry_run: bool,
) -> str:
    """
    Show a two-panel before/after UI.
    Returns: 'y' (apply), 'n' (skip), 'q' (quit)
    """
    console.rule()

    # File header
    console.print(f"\n[bold cyan]FILE:[/bold cyan] {file_path.name}")
    console.print(f"[dim]Path: {file_path}[/dim]")
    console.print(f"[dim]Metadata source: {source}[/dim]\n")

    # Build comparison table
    table = Table(
        box=box.ROUNDED,
        show_header=True,
        header_style="bold",
        expand=True,
        padding=(0, 1),
    )
    table.add_column("Field", style="bold", width=14, no_wrap=True)
    table.add_column("Current", width=40)
    table.add_column("Proposed", width=40)

    has_changes = False
    all_fields = list(dict.fromkeys(list(current.keys()) + list(proposed.keys()) + METADATA_FIELDS))

    for field in all_fields:
        if field not in METADATA_FIELDS:
            continue
        old_val = current.get(field, "")
        new_val = proposed.get(field, "")
        if old_val == new_val and not old_val:
            continue  # Skip empty-on-both sides
        label, old_str, new_str = _field_row(field, old_val, new_val)
        if old_val != new_val:
            has_changes = True
        table.add_row(label, Text.from_markup(old_str), Text.from_markup(new_str))

    console.print(table)

    if not has_changes:
        if dry_run:
            console.print("[dim]No changes detected.[/dim]")
            return "n_nochange"
        console.print("\n[dim]No changes detected.[/dim]")
        console.print("[dim]  s = skip   r = re-search (pick a different match)   q = quit[/dim]")
        choice = Prompt.ask("", choices=["s", "r", "q"], default="s", show_choices=True)
        return {"s": "n_nochange", "r": "research", "q": "q"}.get(choice, "n_nochange")

    if dry_run:
        console.print("[yellow](dry-run mode — no files will be changed)[/yellow]")
        return "y"

    choice = Prompt.ask(
        "\n[bold]Apply changes?[/bold]",
        choices=["y", "n", "q"],
        default="y",
        show_choices=True,
    )
    return choice


# ---------------------------------------------------------------------------
# File renaming
# ---------------------------------------------------------------------------

def _safe_filename_part(text: str) -> str:
    """Strip characters that are unsafe in filenames."""
    # Remove or replace characters not allowed in most filesystems
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def canonical_filename(metadata: dict, original_path: Path) -> str:
    """
    Build a canonical filename like:
      Series Name (Year) #001.cbz
      Series Name (Year) #001 - Issue Title.cbz   (when Title is set)
    Falls back gracefully if fields are missing.
    """
    series = metadata.get("Series", "").strip()
    year   = metadata.get("Year",   "").strip()
    number = metadata.get("Number", "").strip()
    title  = metadata.get("Title",  "").strip()

    if not series:
        return original_path.name

    name = series
    if year:
        name += f" ({year})"
    if number:
        try:
            num_int = int(float(number))
            name += f" #{num_int:03d}"
        except ValueError:
            name += f" #{number}"
    if title:
        name += f" - {_safe_filename_part(title)}"

    ext = original_path.suffix.lower()
    return name + ext


def rename_file(path: Path, metadata: dict, dry_run: bool) -> Path:
    """Rename file to canonical form. Returns new path (or original if unchanged)."""
    new_name = canonical_filename(metadata, path)
    new_path = path.parent / new_name
    if new_path == path:
        return path
    if new_path.exists():
        console.print(f"[yellow]Rename skipped — destination already exists:[/yellow] {new_name}")
        return path
    if not dry_run:
        path.rename(new_path)
        console.print(f"[green]Renamed:[/green] {path.name} → {new_name}")
    else:
        console.print(f"[dim]Would rename:[/dim] {path.name} → {new_name}")
    return new_path


# ---------------------------------------------------------------------------
# Main processing loop
# ---------------------------------------------------------------------------

def find_comics(directory: Path) -> List[Path]:
    """Recursively find all CBZ and CBR files."""
    comics = []
    for ext in ("*.cbz", "*.CBZ", "*.cbr", "*.CBR"):
        comics.extend(directory.rglob(ext))
    comics.sort()
    return comics


def process_file(
    path: Path,
    api_key: Optional[str],
    cache: DiskCache,
    dry_run: bool,
    no_rename: bool,
    no_cache: bool = False,
    auto: bool = False,
    changelog: Optional[List] = None,
) -> str:
    """
    Process a single comic file.
    Returns: 'updated', 'skipped', 'no_change', 'error', 'quit', 'converted'
    """
    is_cbr = path.suffix.lower() == ".cbr"

    # --- CBR: offer conversion ---
    if is_cbr:
        console.print(f"\n[yellow]CBR file detected:[/yellow] {path.name}")
        console.print("[dim]RAR archives can't have metadata rewritten in-place. Convert to CBZ first?[/dim]")
        if not dry_run:
            choice = Prompt.ask("Convert to CBZ?", choices=["y", "n", "q"], default="y")
            if choice == "q":
                return "quit"
            if choice == "n":
                return "skipped"
            cbz_path = convert_cbr_to_cbz(path, dry_run)
            if cbz_path is None:
                return "error"
            path = cbz_path
        else:
            console.print("[dim](dry-run: would convert to CBZ)[/dim]")

    # --- Read existing metadata ---
    current_meta = read_metadata(path)

    # --- Infer from filename ---
    filename_inferred = parse_filename(path)

    # Merge: existing metadata takes priority over filename inference
    inferred = {**filename_inferred, **current_meta}

    # --- Fetch from API (loop allows re-search if user picks 'r') ---
    skip_cache = no_cache
    while True:
        proposed_meta, source = fetch_metadata(inferred, api_key, cache, skip_cache=skip_cache)
        if proposed_meta is None:
            console.print(f"[yellow]Could not find metadata for:[/yellow] {path.name} (series: {inferred.get('Series', '?')})")
            console.print("[dim]Skipping.[/dim]")
            return "skipped"

        # Fill in page count from actual archive
        page_count = get_page_count(path)
        if page_count:
            proposed_meta["PageCount"] = str(page_count)

        # Preserve fields not returned by API that are already set,
        # except fields where a missing API value is meaningful (e.g. Summary, Title —
        # carrying those over from the current file would hide wrong existing data).
        for field in METADATA_FIELDS:
            if field not in proposed_meta and field not in FIELDS_NO_PRESERVE and current_meta.get(field):
                proposed_meta[field] = current_meta[field]

        # --- Confirmation UI ---
        if auto:
            verdict, changes = show_confirmation_ui_auto(path, current_meta, proposed_meta, source)
            if verdict == "no_change":
                return "no_change"
            # fall through to apply; record changes for the end-of-run log
            if changelog is not None and changes:
                changelog.append((path, changes))
            break
        else:
            choice = show_confirmation_ui(path, current_meta, proposed_meta, source, dry_run)
            if choice == "q":
                return "quit"
            if choice == "research":
                skip_cache = True
                continue  # re-fetch bypassing cache, picker will appear again
            if choice in ("n", "n_nochange"):
                return "no_change" if choice == "n_nochange" else "skipped"
            break  # "y" — proceed to apply

    # --- Apply ---
    ok = write_cbz_metadata(path, proposed_meta, dry_run)
    if not ok:
        return "error"

    # --- Rename ---
    if not no_rename:
        rename_file(path, proposed_meta, dry_run)

    return "updated"


def print_summary(stats: dict, errors: List[str], ambiguous: List[str]) -> None:
    console.rule()
    console.print("\n[bold]Summary[/bold]")
    table = Table(box=box.SIMPLE, show_header=False)
    table.add_column("", style="bold")
    table.add_column("")
    table.add_row("Processed", str(stats["processed"]))
    table.add_row("[green]Updated[/green]", str(stats["updated"]))
    table.add_row("[dim]No change[/dim]", str(stats["no_change"]))
    table.add_row("[yellow]Skipped[/yellow]", str(stats["skipped"]))
    table.add_row("[red]Errors[/red]", str(stats["errors"]))
    console.print(table)

    if errors:
        console.print("\n[red]Files with errors:[/red]")
        for e in errors:
            console.print(f"  • {e}")

    if ambiguous:
        console.print("\n[yellow]Files with ambiguous/missing metadata:[/yellow]")
        for a in ambiguous:
            console.print(f"  • {a}")


# ---------------------------------------------------------------------------
# Mode: list (display metadata table)
# ---------------------------------------------------------------------------

# Fields shown in list mode — a concise subset
LIST_FIELDS = ["Series", "Title", "Number", "Volume", "Year", "Publisher", "Writer", "PageCount"]

# Symbols for quick visual health check
_TICK  = "[green]✓[/green]"
_CROSS = "[red]✗[/red]"
# "Core" fields that really should be present for a well-tagged comic
CORE_FIELDS = {"Series", "Number", "Year", "Publisher"}


def run_list_mode(comics: List[Path]) -> None:
    """Display a per-file metadata table followed by a per-series collection summary."""

    # --- Read all metadata up front so we only hit each file once ---
    all_meta: List[Tuple[Path, dict]] = []
    for path in comics:
        meta = read_metadata(path)
        # Fill in anything inferrable from the filename for files with no metadata
        if not meta.get("Series") or not meta.get("Number"):
            inferred = parse_filename(path)
            for k, v in inferred.items():
                if not meta.get(k):
                    meta[k] = v
        all_meta.append((path, meta))

    # -------------------------------------------------------------------------
    # Table 1: per-file detail
    # -------------------------------------------------------------------------
    file_table = Table(
        box=box.ROUNDED,
        show_header=True,
        header_style="bold",
        row_styles=["", "dim"],
        padding=(0, 1),
    )
    file_table.add_column("",          width=1,  no_wrap=True)
    file_table.add_column("File",      min_width=20, max_width=40, no_wrap=True)
    file_table.add_column("Fmt",       width=3,  no_wrap=True)
    file_table.add_column("Series",    min_width=16, max_width=28, no_wrap=True)
    file_table.add_column("#",         width=4,  no_wrap=True)
    file_table.add_column("Vol",       width=3,  no_wrap=True)
    file_table.add_column("Year",      width=4,  no_wrap=True)
    file_table.add_column("Publisher", min_width=10, max_width=18, no_wrap=True)
    file_table.add_column("Writer",    min_width=10, max_width=18, no_wrap=True)
    file_table.add_column("Pages",     width=5,  no_wrap=True)

    missing_core = 0
    for path, meta in all_meta:
        fmt = path.suffix.upper().lstrip(".")
        has_all_core = all(meta.get(f) for f in CORE_FIELDS)
        indicator = _TICK if has_all_core else _CROSS
        if not has_all_core:
            missing_core += 1

        def cell(key: str, max_len: int = 18, _meta: dict = meta) -> str:
            val = _meta.get(key, "")
            if not val:
                return "[dim]—[/dim]"
            return val[:max_len - 1] + "…" if len(val) >= max_len else val

        file_table.add_row(
            indicator,
            path.name[:39] + "…" if len(path.name) > 40 else path.name,
            fmt,
            cell("Series", 28),
            cell("Number", 5),
            cell("Volume", 4),
            meta.get("Year", "") or "[dim]—[/dim]",   # year is always 4 chars, never truncate
            cell("Publisher", 18),
            cell("Writer",    18),
            cell("PageCount", 5),
        )

    console.print(file_table)
    console.print(
        f"\n{len(comics)} file(s)  "
        f"[green]✓ {len(comics) - missing_core} complete[/green]   "
        f"[red]✗ {missing_core} missing core fields[/red] "
        f"[dim](Series / Number / Year / Publisher)[/dim]"
    )

    # -------------------------------------------------------------------------
    # Table 2: per-series collection summary
    # -------------------------------------------------------------------------
    # Group by normalised series name. For each series, track:
    #   - unique issue numbers present (deduplicating variants like English/Klingon)
    #   - the reported total issue count (from Count field)
    #   - publisher (for display)
    from collections import defaultdict

    # series_key → {
    #   "numbers": set of unique normalised issue numbers,
    #   "num_counts": Counter of how many files exist per issue number (to detect variants),
    #   "count": int|None (total from Comic Vine Count field),
    #   "publisher": str, "display": str
    # }
    from collections import Counter
    series_map: dict = defaultdict(lambda: {
        "numbers": set(),
        "num_counts": Counter(),
        "count": None,
        "publisher": "",
        "display": "",
    })

    for path, meta in all_meta:
        series_raw = meta.get("Series", "").strip()
        series_key = series_raw.lower() if series_raw else "__unknown__"
        display     = series_raw or "[dim](no series)[/dim]"

        entry = series_map[series_key]
        if not entry["display"]:
            entry["display"] = display
        if not entry["publisher"] and meta.get("Publisher"):
            entry["publisher"] = meta["Publisher"]

        # Unique issue numbers — normalise to strip leading zeros
        num = meta.get("Number", "").strip()
        if num:
            try:
                norm = str(int(float(num)))
            except ValueError:
                norm = num
            entry["numbers"].add(norm)
            entry["num_counts"][norm] += 1

        # Total issue count from metadata (take highest value seen across files)
        count_str = meta.get("Count", "").strip()
        if count_str:
            try:
                count_int = int(count_str)
                if entry["count"] is None or count_int > entry["count"]:
                    entry["count"] = count_int
            except ValueError:
                pass

    console.print()
    console.rule("[bold]Collection summary by series[/bold]")
    console.print()

    series_table = Table(
        box=box.ROUNDED,
        show_header=True,
        header_style="bold",
        padding=(0, 1),
    )
    series_table.add_column("",            width=1,  no_wrap=True)   # status
    series_table.add_column("Series",      min_width=20, max_width=40, no_wrap=True)
    series_table.add_column("Publisher",   min_width=10, max_width=20, no_wrap=True)
    series_table.add_column("Have",        width=5,  no_wrap=True, justify="right")
    series_table.add_column("Total",       width=5,  no_wrap=True, justify="right")
    series_table.add_column("Missing",     width=7,  no_wrap=True, justify="right")
    series_table.add_column("Notes",       min_width=10)

    any_missing = False
    for series_key in sorted(series_map):
        entry   = series_map[series_key]
        have    = len(entry["numbers"])
        total   = entry["count"]
        pub     = entry["publisher"] or "[dim]—[/dim]"
        display = entry["display"]
        if len(display) > 40:
            display = display[:39] + "…"

        # Count how many issue numbers have more than one file (variants).
        # Comic Vine counts each variant as a separate issue, so we subtract
        # the extra copies to get the adjusted total of distinct story issues.
        variant_groups = sum(
            1 for n, c in entry["num_counts"].items() if c > 1
        )
        extra_files = sum(
            c - 1 for c in entry["num_counts"].values() if c > 1
        )
        adjusted_total = (total - extra_files) if total is not None else None

        # Build a note about variants if any exist
        variant_note = ""
        if variant_groups:
            nums = sorted(
                (n for n, c in entry["num_counts"].items() if c > 1),
                key=lambda x: int(x) if x.isdigit() else 0,
            )
            variant_note = f"[dim]{variant_groups} variant issue(s): #{', #'.join(nums)}[/dim]"

        if adjusted_total is None:
            # No Count data — can't compute missing
            indicator = "[dim]?[/dim]"
            have_str    = str(have)
            total_str   = "[dim]?[/dim]"
            missing_str = "[dim]?[/dim]"
            notes       = "[dim]Run fix first to get series total[/dim]"
        elif have >= adjusted_total:
            indicator   = _TICK
            indicator   = _TICK
            have_str    = f"[green]{have}[/green]"
            total_str   = str(adjusted_total)
            missing_str = "[green]0[/green]"
            notes       = variant_note
        else:
            missing = adjusted_total - have
            any_missing = True
            indicator   = _CROSS
            have_str    = str(have)
            total_str   = str(adjusted_total)
            missing_str = f"[red]{missing}[/red]"
            # List the gaps if the series is small enough to enumerate
            if adjusted_total <= 50:
                all_nums = set(range(1, adjusted_total + 1))
                have_ints = set()
                for n in entry["numbers"]:
                    try:
                        have_ints.add(int(float(n)))
                    except ValueError:
                        pass
                gaps = sorted(all_nums - have_ints)
                notes = _format_gaps(gaps) if gaps else variant_note
            else:
                notes = variant_note

        series_table.add_row(
            indicator, display, pub,
            have_str, total_str, missing_str, notes,
        )

    console.print(series_table)

    if any_missing:
        console.print("\n[red]✗[/red] You are missing issues in one or more series.")
    else:
        console.print("\n[green]✓[/green] Collection appears complete for all series with known totals.")


def _format_gaps(gaps: List[int]) -> str:
    """Format a sorted list of integers as compact ranges, e.g. [1,2,3,7,8] → '#1–3, #7–8'."""
    if not gaps:
        return ""
    ranges = []
    start = end = gaps[0]
    for n in gaps[1:]:
        if n == end + 1:
            end = n
        else:
            ranges.append((start, end))
            start = end = n
    ranges.append((start, end))

    parts = []
    for s, e in ranges:
        parts.append(f"#{s}" if s == e else f"#{s}–{e}")
    result = ", ".join(parts)
    # If the list is very long, truncate
    if len(result) > 60:
        result = result[:57] + "…"
    return f"[dim]missing {result}[/dim]"


# ---------------------------------------------------------------------------
# Mode: convert-cbr (bulk CBR → CBZ without metadata changes)
# ---------------------------------------------------------------------------

def run_convert_cbr_mode(comics: List[Path], dry_run: bool) -> None:
    """Convert all CBR files to CBZ, leaving metadata untouched."""
    cbr_files = [c for c in comics if c.suffix.lower() == ".cbr"]
    if not cbr_files:
        console.print("[yellow]No CBR files found.[/yellow]")
        return

    console.print(f"Found [bold]{len(cbr_files)}[/bold] CBR file(s) to convert.\n")

    converted, skipped, errors = 0, 0, 0
    for i, path in enumerate(cbr_files, 1):
        console.print(f"[dim][{i}/{len(cbr_files)}][/dim] {path.name}")
        cbz_path = path.with_suffix(".cbz")
        if cbz_path.exists():
            console.print(f"  [yellow]Skipped[/yellow] — {cbz_path.name} already exists")
            skipped += 1
            continue
        result = convert_cbr_to_cbz(path, dry_run=dry_run)
        if result:
            converted += 1
        else:
            errors += 1

    console.rule()
    console.print(f"\n[bold]Summary[/bold]")
    table = Table(box=box.SIMPLE, show_header=False)
    table.add_column("", style="bold")
    table.add_column("")
    table.add_row("[green]Converted[/green]", str(converted))
    table.add_row("[yellow]Skipped[/yellow]", str(skipped))
    table.add_row("[red]Errors[/red]",    str(errors))
    console.print(table)


# ---------------------------------------------------------------------------
# Mode: auto (apply all without per-file confirmation)
# ---------------------------------------------------------------------------

def show_confirmation_ui_auto(
    file_path: Path,
    current: dict,
    proposed: dict,
    source: str,
) -> Tuple[str, List[Tuple[str, str, str]]]:
    """
    Auto mode: don't prompt, just compute changes.
    Returns (verdict, changes) where verdict is 'updated'|'no_change'
    and changes is a list of (field, old_val, new_val) tuples.
    """
    changes = []
    for field in METADATA_FIELDS:
        old_val = current.get(field, "")
        new_val = proposed.get(field, "")
        if old_val != new_val:
            changes.append((field, old_val, new_val))
    if changes:
        return "updated", changes
    return "no_change", []


def print_auto_changelog(changelog: List[Tuple[Path, List[Tuple[str, str, str]]]]) -> None:
    """Print a summary of all changes made in auto mode."""
    if not changelog:
        console.print("[dim]No files were changed.[/dim]")
        return

    for path, changes in changelog:
        console.print(f"\n[bold cyan]{path.name}[/bold cyan]")
        table = Table(box=box.SIMPLE, show_header=True, header_style="bold", padding=(0, 1))
        table.add_column("Field", style="bold", width=14)
        table.add_column("Was", width=38)
        table.add_column("Now", width=38)
        for field, old_val, new_val in changes:
            _, old_str, new_str = _field_row(field, old_val, new_val)
            table.add_row(field, Text.from_markup(old_str), Text.from_markup(new_str))
        console.print(table)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Comic book metadata fixer — fetches and applies ComicInfo.xml metadata."
    )
    parser.add_argument("path", help="Comic file (.cbz/.cbr) or directory to scan")

    # Modes (mutually exclusive)
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--list", action="store_true",
        help="Display a metadata summary table for all files (no changes made)",
    )
    mode_group.add_argument(
        "--convert-cbr", action="store_true",
        help="Convert all CBR files to CBZ without modifying metadata",
    )
    mode_group.add_argument(
        "--auto", action="store_true",
        help="Apply all changes without per-file confirmation; print a change log at the end",
    )

    # General options
    parser.add_argument("--dry-run", action="store_true", help="Show changes without writing")
    parser.add_argument("--no-rename", action="store_true", help="Do not rename files")
    parser.add_argument("--no-cache", action="store_true", help="Ignore cached API results and re-fetch")
    parser.add_argument("--api-key", help="Comic Vine API key")
    parser.add_argument("--cache-file", default=str(DEFAULT_CACHE_PATH), help="Path to API cache JSON file")
    args = parser.parse_args()

    target = Path(args.path)
    if not target.exists():
        console.print(f"[red]Path not found:[/red] {target}")
        sys.exit(1)

    # --- Discover files ---
    if target.is_file():
        if target.suffix.lower() not in (".cbz", ".cbr"):
            console.print(f"[red]Not a comic file (expected .cbz or .cbr):[/red] {target.name}")
            sys.exit(1)
        comics = [target]
        console.print(f"\n[bold]Processing:[/bold] {target.name}\n")
    else:
        console.print(f"\n[bold]Scanning:[/bold] {target}")
        comics = find_comics(target)
        if not comics:
            console.print("[yellow]No CBZ or CBR files found.[/yellow]")
            sys.exit(0)
        console.print(f"Found [bold]{len(comics)}[/bold] comic file(s).\n")

    # -------------------------------------------------------------------------
    # Mode: --list
    # -------------------------------------------------------------------------
    if args.list:
        run_list_mode(comics)
        return

    # -------------------------------------------------------------------------
    # Mode: --convert-cbr
    # -------------------------------------------------------------------------
    if args.convert_cbr:
        if args.dry_run:
            console.print("[yellow]DRY RUN MODE — no files will be modified.[/yellow]\n")
        run_convert_cbr_mode(comics, dry_run=args.dry_run)
        return

    # -------------------------------------------------------------------------
    # Normal / --auto mode: metadata fix
    # -------------------------------------------------------------------------

    # API key only needed for metadata modes
    api_key = get_api_key(args.api_key)
    if not api_key:
        api_key = prompt_for_api_key()

    cache = DiskCache(Path(args.cache_file))

    if args.dry_run:
        console.print("[yellow]DRY RUN MODE — no files will be modified.[/yellow]\n")
    if args.auto:
        console.print("[cyan]AUTO MODE — changes will be applied without confirmation.[/cyan]\n")

    stats = {"processed": 0, "updated": 0, "skipped": 0, "no_change": 0, "errors": 0}
    error_files: List[str] = []
    ambiguous_files: List[str] = []
    changelog: List[Tuple[Path, List[Tuple[str, str, str]]]] = []

    for i, comic_path in enumerate(comics, 1):
        if args.auto:
            console.print(f"[dim][{i}/{len(comics)}] {comic_path.name}[/dim]")
        else:
            console.print(f"\n[dim][{i}/{len(comics)}][/dim]")
        stats["processed"] += 1
        try:
            result = process_file(
                comic_path,
                api_key=api_key,
                cache=cache,
                dry_run=args.dry_run,
                no_rename=args.no_rename,
                no_cache=args.no_cache,
                auto=args.auto,
                changelog=changelog,
            )
        except KeyboardInterrupt:
            console.print("\n[yellow]Interrupted by user.[/yellow]")
            break
        except Exception as e:
            console.print(f"[red]Unexpected error processing {comic_path.name}: {e}[/red]")
            result = "error"

        if result == "updated":
            stats["updated"] += 1
        elif result == "skipped":
            stats["skipped"] += 1
            ambiguous_files.append(str(comic_path))
        elif result == "no_change":
            stats["no_change"] += 1
        elif result == "error":
            stats["errors"] += 1
            error_files.append(str(comic_path))
        elif result == "quit":
            console.print("[yellow]Quitting early.[/yellow]")
            break

    # In auto mode, print the change log before the summary
    if args.auto and changelog:
        console.rule()
        console.print(f"\n[bold]Changes applied ({len(changelog)} file(s)):[/bold]")
        print_auto_changelog(changelog)

    print_summary(stats, error_files, ambiguous_files)


if __name__ == "__main__":
    main()
