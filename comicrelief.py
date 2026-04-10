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
import html
import io
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

# Default columns shown in --list mode (in order)
DEFAULT_LIST_FIELDS: List[str] = [
    "Series", "Number", "Volume", "Year", "Publisher", "Writer", "PageCount"
]

# Per-field display config for --list table columns
# Each entry: (header_label, rich_column_kwargs, cell_max_len)
# cell_max_len=None means no truncation (value shown as-is)
FIELD_COLUMN_SPECS: Dict[str, Tuple] = {
    "Series":      ("Series",     {"min_width": 16, "max_width": 28, "no_wrap": True}, 28),
    "Title":       ("Title",      {"min_width": 16, "max_width": 30, "no_wrap": True}, 30),
    "Number":      ("#",          {"width": 4,  "no_wrap": True}, 5),
    "Volume":      ("Vol",        {"width": 3,  "no_wrap": True}, 4),
    "Year":        ("Year",       {"width": 4,  "no_wrap": True}, None),
    "Month":       ("Mo",         {"width": 2,  "no_wrap": True}, 3),
    "Publisher":   ("Publisher",  {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Imprint":     ("Imprint",    {"min_width": 8,  "max_width": 16, "no_wrap": True}, 16),
    "Writer":      ("Writer",     {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Penciller":   ("Penciller",  {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Inker":       ("Inker",      {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Colorist":    ("Colorist",   {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "CoverArtist": ("Cover",      {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Editor":      ("Editor",     {"min_width": 10, "max_width": 18, "no_wrap": True}, 18),
    "Genre":       ("Genre",      {"min_width": 8,  "max_width": 18, "no_wrap": True}, 18),
    "Tags":        ("Tags",       {"min_width": 8,  "max_width": 20, "no_wrap": True}, 20),
    "Characters":  ("Characters", {"min_width": 10, "max_width": 22, "no_wrap": True}, 22),
    "Summary":     ("Summary",    {"min_width": 20, "max_width": 50, "no_wrap": True}, 50),
    "AgeRating":   ("Rating",     {"width": 8,  "no_wrap": True}, 9),
    "Count":       ("Count",      {"width": 5,  "no_wrap": True}, 6),
    "PageCount":   ("Pages",      {"width": 5,  "no_wrap": True}, 5),
    "LanguageISO": ("Lang",       {"width": 4,  "no_wrap": True}, 5),
    "StoryArc":    ("Story Arc",  {"min_width": 10, "max_width": 20, "no_wrap": True}, 20),
    "Format":      ("Format",     {"min_width": 8,  "max_width": 14, "no_wrap": True}, 14),
}

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


def search_comicvine_volumes_all(series_name: str, api_key: str) -> list:
    """Search Comic Vine and return ALL volume results (no scoring, no cache). Used for manual searches."""
    data = _cv_get(
        COMICVINE_SEARCH_URL,
        {
            "query": series_name,
            "resources": "volume",
            "field_list": "id,name,start_year,publisher,count_of_issues,genres",
            "limit": 15,
        },
        api_key,
    )
    if not data or data.get("status_code") != 1:
        return []
    return data.get("results", [])


ENGLISH_PUBLISHERS = {
    "dc comics", "marvel", "image comics", "dark horse comics", "idw publishing",
    "boom! studios", "dynamite entertainment", "vertigo", "wildstorm", "valiant",
    "archie comics", "oni press", "titan comics", "aftershock", "antarctic press",
}


def _score_volume(v: dict, series_name: str, year: Optional[str]) -> int:
    """Score a Comic Vine volume candidate for relevance to a series search."""
    s = 0
    vname = v.get("name", "").lower()
    if vname == series_name.lower():
        s += 10
    elif series_name.lower() in vname:
        s += 5

    # Space-normalised comparison: handles filenames where words are run together,
    # e.g. "startrek alienspotlight" → "startrekalienspotlight"
    #   vs "Star Trek Alien Spotlight" → "startrekalienspotlight"  (exact match!)
    # Also handles hyphenated/dash-separated titles.
    q_nospace = re.sub(r"[\s\-_:]+", "", series_name.lower())
    v_nospace = re.sub(r"[\s\-_:]+", "", vname)
    if q_nospace and q_nospace == v_nospace:
        s += 10
    elif q_nospace and (q_nospace in v_nospace or v_nospace.startswith(q_nospace)):
        s += 5

    if year and v.get("start_year") == year:
        s += 8
    count = v.get("count_of_issues") or 0
    if count > 100:
        s += 4
    elif count > 20:
        s += 2
    elif count > 5:
        s += 1
    pub = v.get("publisher") or {}
    pub_name = (pub.get("name", "") if isinstance(pub, dict) else str(pub)).lower()
    if pub_name in ENGLISH_PUBLISHERS:
        s += 12
    return s


def search_comicvine_volume(series_name: str, year: Optional[str], api_key: str, cache: DiskCache, skip_cache: bool = False) -> Optional[dict]:
    """Search for a volume (series) on Comic Vine. Returns the best matching volume dict."""
    cache_key = f"cv_vol:{series_name.lower()}:{year or ''}"
    if not skip_cache:
        cached = cache.get(cache_key)
        if cached is not None:
            return cached

    data = _cv_get(
        COMICVINE_SEARCH_URL,
        {
            "query": series_name,
            "resources": "volume",
            "field_list": "id,name,start_year,publisher,count_of_issues,description,genres",
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

    best = max(results, key=lambda v: _score_volume(v, series_name, year))
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
                "id,name,issue_number,cover_date,description,deck,page_count,"
                "person_credits,character_credits,location_credits,team_credits,"
                "story_arc_credits,volume"
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


def _approx_matches_inferred(vol: dict, inferred: dict) -> bool:
    """
    Return True if a Comic Vine volume result approximately matches the
    metadata inferred from the current file (existing ComicInfo.xml + filename).

    Scoring:
      +3  space-normalised series name is an exact match
      +2  space-normalised series name is a substring match (≥6 chars)
      +2  start year matches inferred Year
      +2  publisher matches inferred Publisher
    Threshold ≥ 2 → approximate match.
    """
    score = 0
    _STRIP = re.compile(r"[\s\-_:,.()'\"!?]+")

    # Series name comparison
    inferred_series = (inferred.get("Series") or "").lower()
    if inferred_series:
        q = _STRIP.sub("", inferred_series)
        v = _STRIP.sub("", (vol.get("name") or "").lower())
        if q and q == v:
            score += 3
        elif q and len(q) >= 6 and (q in v or v.startswith(q)):
            score += 2

    # Year match
    inferred_year = inferred.get("Year")
    if inferred_year and str(vol.get("start_year", "")) == str(inferred_year):
        score += 2

    # Publisher match (only meaningful if existing metadata already has it)
    inferred_pub = _STRIP.sub("", (inferred.get("Publisher") or "").lower())
    vol_pub_obj = vol.get("publisher") or {}
    vol_pub = _STRIP.sub(
        "",
        (vol_pub_obj.get("name", "") if isinstance(vol_pub_obj, dict) else str(vol_pub_obj)).lower(),
    )
    if inferred_pub and vol_pub and (inferred_pub in vol_pub or vol_pub in inferred_pub):
        score += 2

    return score >= 2


def _pick_volume(
    results: list,
    highlight_count: Optional[int] = None,
    inferred: Optional[dict] = None,
) -> dict:
    """Show a numbered list of Comic Vine volumes and let the user choose one.

    highlight_count  — when set, rows whose issue count matches are highlighted
                       in green (★) as a likely collection size match.
    inferred         — when set, rows that approximately match the file's own
                       metadata are flagged with a red * (metadata hint).
    """
    legends = []
    if highlight_count:
        legends.append(f"[bold green]★[/bold green] = {highlight_count} issues matches your folder")
    if inferred:
        legends.append("[red]*[/red] = series name / year / publisher matches file metadata")
    if legends:
        console.print("  [dim]" + "   ".join(legends) + "[/dim]")

    console.print(f"\n[yellow]Found {len(results)} series — please choose:[/yellow]")
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("#",         width=4,  no_wrap=True)
    table.add_column("",          width=1,  no_wrap=True)   # metadata-match indicator
    table.add_column("Series",    min_width=24, no_wrap=True)
    table.add_column("Start",     width=6,  no_wrap=True)
    table.add_column("Publisher", min_width=14, no_wrap=True)
    table.add_column("Issues",    width=7,  no_wrap=True, justify="right")
    table.add_column("ID",        width=8,  no_wrap=True)

    for i, vol in enumerate(results, 1):
        pub = vol.get("publisher") or {}
        pub_name = pub.get("name", "—") if isinstance(pub, dict) else str(pub)
        count = vol.get("count_of_issues")
        count_str = str(count) if count is not None else "?"

        count_match = highlight_count is not None and count == highlight_count
        meta_match  = inferred is not None and _approx_matches_inferred(vol, inferred)

        issues_cell = f"[bold green]{count_str} ★[/bold green]" if count_match else count_str
        meta_cell   = "[red]*[/red]" if meta_match else ""

        if count_match:
            # Highlight the whole row in green for the count match
            table.add_row(
                f"[bold green]{i}[/bold green]",
                meta_cell,
                f"[bold green]{vol.get('name') or '?'}[/bold green]",
                f"[bold green]{vol.get('start_year') or '?'}[/bold green]",
                f"[bold green]{pub_name}[/bold green]",
                issues_cell,
                f"[bold green]{vol.get('id', '')}[/bold green]",
            )
        else:
            table.add_row(
                str(i),
                meta_cell,
                vol.get("name") or "?",
                str(vol.get("start_year") or "?"),
                pub_name,
                issues_cell,
                str(vol.get("id", "")),
            )
    console.print(table)

    choices = [str(i) for i in range(1, len(results) + 1)]
    choice = Prompt.ask("Enter number", choices=choices, default="1")
    return results[int(choice) - 1]


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


def _clean_html(raw: str) -> str:
    """Strip HTML from a Comic Vine description string and return plain text."""
    raw = re.sub(r"<table\b[^>]*>.*?</table>", "", raw, flags=re.IGNORECASE | re.DOTALL)
    raw = re.sub(r"</?(p|div|h[1-6]|ul|ol|li|blockquote|section)\b[^>]*>", "\n\n", raw, flags=re.IGNORECASE)
    raw = re.sub(r"<br\s*/?>", "\n", raw, flags=re.IGNORECASE)
    raw = re.sub(r"<[^>]+>", "", raw)
    raw = html.unescape(raw)
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n[ \t]+", "\n", raw)
    raw = re.sub(r"\n{3,}", "\n\n", raw)
    return raw.strip()


def extract_cv_metadata(volume: dict, issue: Optional[dict], full_metadata: bool = False) -> dict:
    """Convert Comic Vine volume + issue dicts into our flat metadata format.

    When full_metadata is True, additional fields are populated:
    - LanguageISO (always "en" for Comic Vine data)
    - Genre (from the volume's genre list)
    - Tags (from the issue's location and team credits)
    - Characters cap is removed (all characters included)
    - Summary falls back to the issue deck, then the volume description
    """
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

        # Characters (capped at 20 normally; unlimited with full_metadata)
        chars = [c.get("name", "") for c in issue.get("character_credits", [])]
        if chars:
            meta["Characters"] = ", ".join(chars if full_metadata else chars[:20])

        # Story arcs
        arcs = [a.get("name", "") for a in issue.get("story_arc_credits", [])]
        if arcs:
            meta["StoryArc"] = ", ".join(arcs)

        # CV page count — stashed under a private key so process_file can
        # compare it against the actual archive count without writing it to
        # ComicInfo.xml (PageCount is always set from the real archive count).
        cv_pc = issue.get("page_count")
        if cv_pc is not None:
            try:
                meta["_cv_page_count"] = int(cv_pc)
            except (TypeError, ValueError):
                pass

        # Summary — clean HTML description
        desc = _clean_html(issue.get("description", "") or "")
        if desc:
            meta["Summary"] = desc[:2000]

        if full_metadata:
            # Tags: locations + teams from this issue
            locations = [loc.get("name", "") for loc in issue.get("location_credits", []) if loc.get("name")]
            teams     = [t.get("name", "")   for t   in issue.get("team_credits", [])     if t.get("name")]
            tag_items = locations + teams
            if tag_items:
                meta["Tags"] = ", ".join(tag_items)

            # Summary fallback 1: issue deck (short tagline)
            if not meta.get("Summary"):
                deck = (issue.get("deck") or "").strip()
                if deck:
                    meta["Summary"] = html.unescape(deck)

    if full_metadata:
        # Language — CV is an English-language database
        meta["LanguageISO"] = "en"

        # Genre from the volume's genre list
        genres = volume.get("genres") or []
        if isinstance(genres, list) and genres:
            genre_names = [g.get("name", "") for g in genres if isinstance(g, dict) and g.get("name")]
            if genre_names:
                meta["Genre"] = ", ".join(genre_names)

        # Summary fallback 2: volume description (series overview)
        if not meta.get("Summary"):
            vol_desc = _clean_html(volume.get("description", "") or "")
            if vol_desc:
                meta["Summary"] = vol_desc[:2000]

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
# GCD (Grand Comics Database) API
# ---------------------------------------------------------------------------

GCD_BASE_URL = "https://www.comics.org/api"
GCD_RATE_LIMIT = 0.5  # seconds between requests
_last_gcd_request = 0.0


def _gcd_get(endpoint: str, params: dict) -> Optional[dict]:
    global _last_gcd_request
    elapsed = time.time() - _last_gcd_request
    if elapsed < GCD_RATE_LIMIT:
        time.sleep(GCD_RATE_LIMIT - elapsed)
    url = f"{GCD_BASE_URL}/{endpoint}"
    try:
        r = requests.get(
            url,
            params={**params, "format": "json"},
            timeout=15,
            headers={"User-Agent": "comicrelief/1.0"},
        )
        _last_gcd_request = time.time()
        if r.status_code == 200:
            return r.json()
    except requests.RequestException:
        pass
    return None


def search_gcd(
    series_name: str,
    issue_number: Optional[str],
    year: Optional[str],
    cache: DiskCache,
) -> Optional[dict]:
    """Search GCD for a series+issue and return flat metadata dict."""
    cache_key = f"gcd:{series_name.lower()}:{issue_number or ''}:{year or ''}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    # --- 1. Find the best matching series ---
    data = _gcd_get("series/", {"name": series_name})
    if not data or not data.get("results"):
        cache.set(cache_key, None)
        return None

    results = data["results"]

    def _score_gcd_series(s: dict) -> int:
        sc = 0
        sname = (s.get("name") or "").lower()
        if sname == series_name.lower():
            sc += 10
        elif series_name.lower() in sname:
            sc += 5
        yr = s.get("year_began")
        if year and yr and str(yr) == str(year):
            sc += 8
        count = s.get("issue_count") or 0
        if count > 50:
            sc += 3
        elif count > 10:
            sc += 1
        lang = s.get("language") or {}
        if isinstance(lang, dict) and lang.get("name", "").lower() in ("english", "en"):
            sc += 4
        return sc

    best_series = max(results, key=_score_gcd_series)
    series_id = best_series.get("id")
    if not series_id:
        cache.set(cache_key, None)
        return None

    pub = best_series.get("publisher") or {}
    meta: dict = {
        "Series": best_series.get("name", ""),
    }
    if isinstance(pub, dict) and pub.get("name"):
        meta["Publisher"] = pub["name"]
    elif isinstance(pub, str) and pub:
        meta["Publisher"] = pub

    if not issue_number:
        result = {k: v for k, v in meta.items() if v}
        cache.set(cache_key, result or None)
        return result or None

    # --- 2. Find the specific issue ---
    issue_data = _gcd_get("issue/", {"series": series_id, "number": issue_number})
    if not issue_data or not issue_data.get("results"):
        # Return series-level data rather than nothing
        result = {k: v for k, v in meta.items() if v}
        cache.set(cache_key, result or None)
        return result or None

    issues = issue_data["results"]
    target_norm = str(issue_number).lstrip("0") or "0"

    def _issue_score(i: dict) -> int:
        n = str(i.get("number", "")).strip().lstrip("0") or "0"
        return 1 if n == target_norm else 0

    issue = max(issues, key=_issue_score)

    # Publication date
    pub_date = issue.get("publication_date") or issue.get("on_sale_date") or ""
    if pub_date:
        yr_m = re.search(r"\b(\d{4})\b", pub_date)
        if yr_m:
            meta["Year"] = yr_m.group(1)
        month_names = {
            "january": "1", "february": "2", "march": "3", "april": "4",
            "may": "5", "june": "6", "july": "7", "august": "8",
            "september": "9", "october": "10", "november": "11", "december": "12",
        }
        mo_m = re.search(
            r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b",
            pub_date,
            re.IGNORECASE,
        )
        if mo_m:
            meta["Month"] = month_names.get(mo_m.group(1).lower(), "")

    issue_title = (issue.get("title") or "").strip()
    if issue_title:
        meta["Title"] = issue_title

    # --- 3. Fetch full issue detail for credits ---
    issue_id = issue.get("id")
    if issue_id:
        detail = _gcd_get(f"issue/{issue_id}/", {})
        if detail:
            # story_set contains sequences; each has a credits list
            story_set = detail.get("story_set") or []
            writers: List[str] = []
            pencillers: List[str] = []
            inkers: List[str] = []
            colorists: List[str] = []
            editors: List[str] = []

            for story in story_set:
                # Skip cover-only sequences if there are story sequences
                for credit in story.get("credits") or []:
                    role_obj = credit.get("role") or {}
                    role = (role_obj.get("name", "") if isinstance(role_obj, dict) else str(role_obj)).lower()
                    person_obj = credit.get("person") or {}
                    name = (person_obj.get("name", "") if isinstance(person_obj, dict) else str(person_obj)).strip()
                    if not name:
                        continue
                    if "script" in role or "writer" in role:
                        writers.append(name)
                    elif "pencil" in role:
                        pencillers.append(name)
                    elif "ink" in role:
                        inkers.append(name)
                    elif "color" in role or "colour" in role:
                        colorists.append(name)
                    elif "edit" in role:
                        editors.append(name)

            # Deduplicate while preserving order
            def _dedup(lst: list) -> list:
                seen: set = set()
                return [x for x in lst if not (x in seen or seen.add(x))]  # type: ignore[func-returns-value]

            if writers:
                meta["Writer"] = ", ".join(_dedup(writers))
            if pencillers:
                meta["Penciller"] = ", ".join(_dedup(pencillers))
            if inkers:
                meta["Inker"] = ", ".join(_dedup(inkers))
            if colorists:
                meta["Colorist"] = ", ".join(_dedup(colorists))
            if editors:
                meta["Editor"] = ", ".join(_dedup(editors))

    result = {k: v for k, v in meta.items() if v}
    cache.set(cache_key, result or None)
    return result or None


# ---------------------------------------------------------------------------
# MangaDex API (manga-specific)
# ---------------------------------------------------------------------------

MANGADEX_BASE_URL = "https://api.mangadex.org"
MANGADEX_RATE_LIMIT = 0.3  # ~3 req/s (well within their 5 req/s limit)
_last_mangadex_request = 0.0


def _mangadex_get(endpoint: str, params: list) -> Optional[dict]:
    """params should be a list of (key, value) tuples to support repeated keys."""
    global _last_mangadex_request
    elapsed = time.time() - _last_mangadex_request
    if elapsed < MANGADEX_RATE_LIMIT:
        time.sleep(MANGADEX_RATE_LIMIT - elapsed)
    url = f"{MANGADEX_BASE_URL}/{endpoint}"
    try:
        r = requests.get(
            url,
            params=params,
            timeout=15,
            headers={"User-Agent": "comicrelief/1.0"},
        )
        _last_mangadex_request = time.time()
        if r.status_code == 200:
            return r.json()
    except requests.RequestException:
        pass
    return None


def search_mangadex(
    series_name: str,
    issue_number: Optional[str],
    year: Optional[str],
    cache: DiskCache,
) -> Optional[dict]:
    """Search MangaDex for a manga series and return flat metadata dict."""
    cache_key = f"mangadex:{series_name.lower()}:{issue_number or ''}:{year or ''}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    data = _mangadex_get(
        "manga",
        [
            ("title", series_name),
            ("includes[]", "author"),
            ("includes[]", "artist"),
            ("limit", "5"),
            ("order[relevance]", "desc"),
        ],
    )
    if not data or not data.get("data"):
        cache.set(cache_key, None)
        return None

    results = data["data"]
    if not results:
        cache.set(cache_key, None)
        return None

    def _score_md(m: dict) -> int:
        attrs = m.get("attributes") or {}
        sc = 0
        for t in (attrs.get("title") or {}).values():
            if t.lower() == series_name.lower():
                sc += 10
                break
            if series_name.lower() in t.lower():
                sc += 5
                break
        if year and str(attrs.get("year") or "") == str(year):
            sc += 8
        return sc

    best = max(results, key=_score_md)
    attrs = best.get("attributes") or {}

    # Prefer English title, fall back to romanised or first available
    titles = attrs.get("title") or {}
    title_en = (
        titles.get("en")
        or titles.get("ja-ro")
        or next(iter(titles.values()), "")
    )

    meta: dict = {}
    if title_en:
        meta["Series"] = title_en.strip()

    yr = attrs.get("year")
    if yr:
        meta["Year"] = str(yr)

    # Description: English preferred, any language as fallback
    descs = attrs.get("description") or {}
    desc_raw = descs.get("en", "") or next(iter(descs.values()), "") if descs else ""
    if desc_raw:
        # Strip basic Markdown (MangaDex uses CommonMark)
        desc_clean = re.sub(r"\[([^\]]+)\]\([^\)]+\)", r"\1", desc_raw)  # links
        desc_clean = re.sub(r"[*_]{1,2}([^*_]+)[*_]{1,2}", r"\1", desc_clean)  # bold/italic
        desc_clean = re.sub(r"^#+\s*", "", desc_clean, flags=re.MULTILINE)  # headers
        desc_clean = re.sub(r"\n{3,}", "\n\n", desc_clean)
        meta["Summary"] = desc_clean.strip()[:2000]

    # Tags → Genre + Tags
    genre_tags: List[str] = []
    content_tags: List[str] = []
    for tag in (attrs.get("tags") or []):
        tag_attrs = tag.get("attributes") or {}
        tag_name = (tag_attrs.get("name") or {}).get("en", "")
        group = tag_attrs.get("group", "")
        if not tag_name:
            continue
        if group == "genre":
            genre_tags.append(tag_name)
        else:
            content_tags.append(tag_name)
    if genre_tags:
        meta["Genre"] = ", ".join(genre_tags)
    if content_tags:
        meta["Tags"] = ", ".join(content_tags[:10])

    meta["LanguageISO"] = "ja"

    # Author/Artist from relationships
    authors: List[str] = []
    artists: List[str] = []
    for rel in (best.get("relationships") or []):
        rel_type = rel.get("type", "")
        rel_attrs = rel.get("attributes") or {}
        name = (rel_attrs.get("name") or "").strip()
        if not name:
            continue
        if rel_type == "author":
            authors.append(name)
        elif rel_type == "artist":
            artists.append(name)
    if authors:
        meta["Writer"] = ", ".join(authors)
    if artists:
        meta["Penciller"] = ", ".join(artists)

    result = {k: v for k, v in meta.items() if v}
    cache.set(cache_key, result or None)
    return result or None


# ---------------------------------------------------------------------------
# AniList API (GraphQL, manga series-level)
# ---------------------------------------------------------------------------

ANILIST_URL = "https://graphql.anilist.co"
ANILIST_RATE_LIMIT = 2.1  # 30 req/min → 2s apart is comfortably under the limit
_last_anilist_request = 0.0

_ANILIST_QUERY = """
query ($search: String) {
  Media(search: $search, type: MANGA) {
    title { romaji english native }
    description(asHtml: false)
    genres
    startDate { year }
    staff(perPage: 6) {
      edges {
        role
        node { name { full } }
      }
    }
  }
}
"""


def _anilist_post(variables: dict) -> Optional[dict]:
    global _last_anilist_request
    elapsed = time.time() - _last_anilist_request
    if elapsed < ANILIST_RATE_LIMIT:
        time.sleep(ANILIST_RATE_LIMIT - elapsed)
    try:
        r = requests.post(
            ANILIST_URL,
            json={"query": _ANILIST_QUERY, "variables": variables},
            headers={"Content-Type": "application/json", "User-Agent": "comicrelief/1.0"},
            timeout=15,
        )
        _last_anilist_request = time.time()
        if r.status_code == 200:
            return r.json()
    except requests.RequestException:
        pass
    return None


def search_anilist(series_name: str, cache: DiskCache) -> Optional[dict]:
    """Search AniList for a manga series and return flat metadata dict (series-level only)."""
    cache_key = f"anilist:{series_name.lower()}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    data = _anilist_post({"search": series_name})
    if not data:
        cache.set(cache_key, None)
        return None

    media = (data.get("data") or {}).get("Media")
    if not media:
        cache.set(cache_key, None)
        return None

    meta: dict = {}

    # Title: English preferred, romanised as fallback
    titles = media.get("title") or {}
    title_en = titles.get("english") or titles.get("romaji") or ""
    if title_en:
        meta["Series"] = title_en.strip()

    # Year
    start = media.get("startDate") or {}
    if start.get("year"):
        meta["Year"] = str(start["year"])

    # Description (asHtml: false, but may still have <br> tags)
    desc = (media.get("description") or "").strip()
    if desc:
        desc = re.sub(r"<br\s*/?>", "\n", desc, flags=re.IGNORECASE)
        desc = re.sub(r"<[^>]+>", "", desc)
        desc = html.unescape(desc)
        desc = re.sub(r"\n{3,}", "\n\n", desc).strip()
        if desc:
            meta["Summary"] = desc[:2000]

    # Genres
    genres = media.get("genres") or []
    if genres:
        meta["Genre"] = ", ".join(genres)

    # Staff credits
    staff_edges = (media.get("staff") or {}).get("edges") or []
    writers: List[str] = []
    artists: List[str] = []
    for edge in staff_edges:
        role = (edge.get("role") or "").lower()
        node = edge.get("node") or {}
        name = ((node.get("name") or {}).get("full") or "").strip()
        if not name:
            continue
        if any(kw in role for kw in ("story", "script", "author", "original")):
            writers.append(name)
        elif any(kw in role for kw in ("art", "illustrat", "character design", "draw")):
            artists.append(name)
    if writers:
        meta["Writer"] = ", ".join(writers)
    if artists:
        meta["Penciller"] = ", ".join(artists)

    result = {k: v for k, v in meta.items() if v}
    cache.set(cache_key, result or None)
    return result or None


# ---------------------------------------------------------------------------
# Supplemental metadata merging
# ---------------------------------------------------------------------------

MANGA_PUBLISHERS: set = {
    "viz media", "viz", "kodansha", "kodansha comics", "kodansha usa",
    "yen press", "seven seas entertainment", "seven seas",
    "dark horse manga", "tokyopop", "del rey manga", "vertical",
    "square enix manga", "j-novel club", "one peace books",
    "shueisha", "kadokawa", "hakusensha", "shogakukan",
}


def _is_manga(meta: dict) -> bool:
    """Detect if a comic is likely manga from publisher or language."""
    pub = (meta.get("Publisher") or "").lower()
    if any(mp in pub for mp in MANGA_PUBLISHERS):
        return True
    return (meta.get("LanguageISO") or "").lower() in ("ja", "ko", "zh")


def _needs_supplement(meta: dict) -> bool:
    """Return True if key fields are missing that supplemental sources might fill."""
    return not meta.get("Summary") or not meta.get("Writer")


def _merge_metadata(*sources: Optional[dict]) -> dict:
    """
    Merge metadata dicts from multiple sources in priority order.
    For each field, use the first non-empty value across sources.
    The first source (primary) is never overwritten.
    """
    merged: dict = {}
    for source in sources:
        if not source:
            continue
        for field in METADATA_FIELDS:
            if field not in merged and source.get(field):
                merged[field] = source[field]
    return merged


def _supplement_metadata(
    meta: dict,
    search_name: str,
    issue_num: Optional[str],
    year: Optional[str],
    full_metadata: bool,
    cache: DiskCache,
) -> Tuple[dict, List[str]]:
    """
    Try supplemental sources (GCD, MangaDex, AniList) to fill in missing fields.
    Returns (merged_metadata, list_of_supplemental_source_names_used).
    """
    if not _needs_supplement(meta):
        return meta, []

    is_manga_series = _is_manga(meta)
    supp_metas: List[Optional[dict]] = []
    supp_names: List[str] = []

    if is_manga_series:
        md_meta = search_mangadex(search_name, issue_num, year, cache)
        if md_meta:
            supp_metas.append(md_meta)
            supp_names.append("MangaDex")
        # Only call AniList if we still need data after MangaDex
        still_needed = _needs_supplement(_merge_metadata(meta, md_meta) if md_meta else meta)
        if still_needed:
            al_meta = search_anilist(search_name, cache)
            if al_meta:
                supp_metas.append(al_meta)
                supp_names.append("AniList")
    else:
        gcd_meta = search_gcd(search_name, issue_num, year, cache)
        if gcd_meta:
            supp_metas.append(gcd_meta)
            supp_names.append("GCD")

    if supp_metas:
        return _merge_metadata(meta, *supp_metas), supp_names
    return meta, []


# ---------------------------------------------------------------------------
# Smart cover matching (perceptual hash)
# ---------------------------------------------------------------------------

def extract_cover_image(path: Path) -> Optional[bytes]:
    """
    Extract the first image from a CBZ archive (the cover page).
    Returns raw image bytes, or None if extraction fails or format is not CBZ.
    """
    if path.suffix.lower() != ".cbz":
        return None
    try:
        with zipfile.ZipFile(path) as zf:
            image_exts = ('.jpg', '.jpeg', '.png', '.webp', '.gif', '.tiff', '.bmp')
            image_names = sorted([
                n for n in zf.namelist()
                if n.lower().endswith(image_exts)
                and '__MACOSX' not in n
                and not os.path.basename(n).startswith('.')
            ])
            if image_names:
                return zf.read(image_names[0])
    except Exception:
        pass
    return None


def _compute_phash(image_bytes: bytes) -> Optional[object]:
    """
    Compute a perceptual hash of image bytes.
    Returns an imagehash.ImageHash, or None if PIL/imagehash are unavailable
    or the image cannot be decoded.
    """
    try:
        from PIL import Image
        import imagehash
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        return imagehash.phash(img)
    except Exception:
        return None


def _download_phash(url: str) -> Optional[object]:
    """Download an image from a URL and return its perceptual hash, or None on failure."""
    try:
        resp = requests.get(url, timeout=15, headers={"User-Agent": "comicrelief/1.0"})
        if resp.status_code == 200:
            return _compute_phash(resp.content)
    except Exception:
        pass
    return None


def _get_issue_cover_url(volume_id: int, issue_number: str, api_key: str) -> Optional[str]:
    """
    Fetch a Comic Vine issue's cover image URL using a lightweight API call
    (only requests id, issue_number, and image fields). Used during smart matching.
    Returns the medium_url (or small_url as fallback), or None.
    """
    normalized = (
        str(int(float(issue_number)))
        if issue_number.replace(".", "").isdigit()
        else issue_number
    )
    data = _cv_get(
        COMICVINE_ISSUES_URL,
        {
            "filter": f"volume:{volume_id},issue_number:{normalized}",
            "field_list": "id,issue_number,image",
            "limit": 1,
        },
        api_key,
    )
    if not data or data.get("status_code") != 1:
        return None
    results = data.get("results", [])
    if not results:
        return None
    image = results[0].get("image") or {}
    return image.get("medium_url") or image.get("small_url")


def _get_cv_candidates(
    series_name: str,
    year: Optional[str],
    api_key: str,
    cache: "DiskCache",
    skip_cache: bool = False,
    n: int = 5,
) -> list:
    """
    Search Comic Vine and return the top-N scored candidate volumes.
    Uses its own cache key so it doesn't interfere with the single-best cache
    used by search_comicvine_volume.
    """
    cache_key = f"cv_vol_list:{series_name.lower()}:{year or ''}"
    if not skip_cache:
        cached = cache.get(cache_key)
        if cached is not None:
            return cached

    data = _cv_get(
        COMICVINE_SEARCH_URL,
        {
            "query": series_name,
            "resources": "volume",
            "field_list": "id,name,start_year,publisher,count_of_issues,genres",
            "limit": 15,
        },
        api_key,
    )
    if not data or data.get("status_code") != 1:
        cache.set(cache_key, [])
        return []

    results = data.get("results", [])

    # If the primary query returned nothing, retry with individual tokens.
    # Handles filenames like "startrek_alienspotlight" where neither token is a
    # real word that CV's full-text search can match — but "alienspotlight" as a
    # query might surface volumes containing "Alien Spotlight", etc.
    if not results:
        seen_ids: set = set()
        tokens = [t for t in series_name.split() if len(t) >= 5]
        for token in tokens[:3]:   # cap retries to avoid burning rate-limit budget
            retry_data = _cv_get(
                COMICVINE_SEARCH_URL,
                {
                    "query": token,
                    "resources": "volume",
                    "field_list": "id,name,start_year,publisher,count_of_issues,genres",
                    "limit": 10,
                },
                api_key,
            )
            for v in (retry_data or {}).get("results", []):
                vid = v.get("id")
                if vid and vid not in seen_ids:
                    results.append(v)
                    seen_ids.add(vid)

    if not results:
        cache.set(cache_key, [])
        return []

    scored = sorted(results, key=lambda v: _score_volume(v, series_name, year), reverse=True)
    top = scored[:n]
    cache.set(cache_key, top)
    return top


def smart_match_volume(
    candidates: list,
    issue_number: Optional[str],
    cover_bytes: Optional[bytes],
    api_key: str,
) -> Optional[dict]:
    """
    Given multiple candidate Comic Vine volumes and the comic's cover image bytes,
    download the issue cover from each candidate and use perceptual hashing to
    pick the visually closest match.

    Returns the best-matching volume dict, or None if matching is inconclusive
    (PIL/imagehash unavailable, no covers found, or all distances too high).
    """
    if not cover_bytes or not issue_number or len(candidates) < 2:
        return None

    cover_hash = _compute_phash(cover_bytes)
    if cover_hash is None:
        # PIL/imagehash not installed, or cover image unreadable — degrade gracefully
        return None

    console.print(f"  [dim]Smart matching: comparing cover image against {len(candidates)} candidates…[/dim]")

    MATCH_THRESHOLD = 40  # pHash distances: 0=identical, 64=max. >40 = probably no match.
    CONFIDENT_THRESHOLD = 15  # < 15 = high confidence

    best_vol: Optional[dict] = None
    best_distance: int = 999
    results_log: list = []

    for vol in candidates:
        vol_id = vol.get("id")
        if not vol_id:
            continue
        pub = (vol.get("publisher") or {}).get("name", "?")
        label = f"{vol.get('name', '?')} ({vol.get('start_year', '?')}, {pub})"

        cover_url = _get_issue_cover_url(vol_id, issue_number, api_key)
        if not cover_url:
            results_log.append(f"    [dim]{label}: no cover available[/dim]")
            continue

        candidate_hash = _download_phash(cover_url)
        if candidate_hash is None:
            results_log.append(f"    [dim]{label}: could not fetch cover[/dim]")
            continue

        distance = cover_hash - candidate_hash  # imagehash hamming distance
        similarity_pct = max(0, 100 - int(distance / 64.0 * 100))
        results_log.append(
            f"    [dim]{label}: {similarity_pct}% match (distance={distance})[/dim]"
        )

        if distance < best_distance:
            best_distance = distance
            best_vol = vol

    for line in results_log:
        console.print(line)

    if best_vol is None or best_distance > MATCH_THRESHOLD:
        console.print("  [dim]Smart match inconclusive — using score-based result.[/dim]")
        return None

    pub = (best_vol.get("publisher") or {}).get("name", "?")
    confidence = "high confidence" if best_distance <= CONFIDENT_THRESHOLD else "best available"
    console.print(
        f"  [green]Smart match ({confidence}):[/green] "
        f"{best_vol.get('name')} ({best_vol.get('start_year')}, {pub}) "
        f"— distance {best_distance}/64"
    )
    return best_vol


# ---------------------------------------------------------------------------
# Metadata lookup orchestration
# ---------------------------------------------------------------------------

def fetch_comicvine_volume_by_id(volume_id: int, api_key: str, cache: DiskCache) -> Optional[dict]:
    """
    Fetch a Comic Vine volume directly by its numeric ID.
    The ID is the number at the end of the Comic Vine URL:
      https://comicvine.gamespot.com/series-name/4050-<ID>/
    """
    cache_key = f"cv_vol_id:{volume_id}"
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    field_list = "id,name,start_year,publisher,count_of_issues,description"

    # Primary: single-resource endpoint
    data = _cv_get(
        f"https://comicvine.gamespot.com/api/volume/4050-{volume_id}/",
        {"field_list": field_list},
        api_key,
    )
    if data and data.get("status_code") == 1:
        result = data.get("results")
        if result and isinstance(result, dict):
            cache.set(cache_key, result)
            return result

    # Fallback: list endpoint with id filter
    data = _cv_get(
        COMICVINE_VOLUMES_URL,
        {"filter": f"id:{volume_id}", "field_list": field_list, "limit": 1},
        api_key,
    )
    if data and data.get("status_code") == 1:
        results = data.get("results", [])
        if results:
            cache.set(cache_key, results[0])
            return results[0]

    return None


def fetch_metadata(
    inferred: dict,
    api_key: Optional[str],
    cache: "DiskCache",
    skip_cache: bool = False,
    volume_override: Optional[dict] = None,  # pre-selected Comic Vine volume dict
    cover_bytes: Optional[bytes] = None,      # raw cover image for smart matching
    full_metadata: bool = False,              # fetch and store all available fields
) -> Tuple[Optional[dict], str]:
    """
    Fetch metadata from Comic Vine (primary) or Metron (fallback).
    After the primary source, supplemental sources (GCD / MangaDex / AniList)
    are queried to fill in any missing Summary or Writer fields.

    When cover_bytes is provided and multiple candidate series exist, smart cover
    matching is used to pick the best series match via perceptual hash comparison.

    Returns (metadata_dict, source_label) or (None, "not found").
    """
    issue_num = inferred.get("Number")
    series = inferred.get("Series", "")
    search_name = slugify_series(series) if series else ""
    year = inferred.get("Year")

    # --- Use override volume if provided (user already picked the right series) ---
    if volume_override and api_key:
        vol_id = volume_override.get("id")
        issue = None
        if vol_id and issue_num:
            issue = fetch_comicvine_issue(vol_id, issue_num, api_key, cache, skip_cache=skip_cache)
        meta = extract_cv_metadata(volume_override, issue, full_metadata=full_metadata)
        if search_name:
            meta, supp = _supplement_metadata(meta, search_name, issue_num, year, full_metadata, cache)
            source = "Comic Vine" + (f" + {', '.join(supp)}" if supp else "")
        else:
            source = "Comic Vine"
        return meta, source

    if not series:
        return None, "no series name"

    # --- Comic Vine ---
    if api_key:
        # Get the top-N scored candidates for smart matching
        candidates = _get_cv_candidates(search_name, year, api_key, cache, skip_cache=skip_cache)

        volume: Optional[dict] = None

        if len(candidates) >= 2 and cover_bytes and issue_num:
            # Smart matching: compare cover image against each candidate's issue cover
            volume = smart_match_volume(candidates, issue_num, cover_bytes, api_key)

        if volume is None:
            # Use the top score-based candidate
            volume = candidates[0] if candidates else None

        if volume:
            vol_id = volume.get("id")
            issue = None
            if vol_id and issue_num:
                issue = fetch_comicvine_issue(vol_id, issue_num, api_key, cache, skip_cache=skip_cache)
            meta = extract_cv_metadata(volume, issue, full_metadata=full_metadata)
            cv_source = "Comic Vine (smart match)" if len(candidates) >= 2 and cover_bytes else "Comic Vine"
            meta, supp = _supplement_metadata(meta, search_name, issue_num, year, full_metadata, cache)
            source = cv_source + (f" + {', '.join(supp)}" if supp else "")
            return meta, source

    # --- Metron fallback ---
    meta = search_metron(search_name, issue_num, year, cache)
    if meta:
        meta, supp = _supplement_metadata(meta, search_name, issue_num, year, full_metadata, cache)
        source = "Metron" + (f" + {', '.join(supp)}" if supp else "")
        return meta, source

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
        console.print(
            "[dim]  s = skip   r = re-search (bypass cache)   "
            "n = search new series name   i = enter Comic Vine volume ID   q = quit[/dim]"
        )
        choice = Prompt.ask("", choices=["s", "r", "n", "i", "q"], default="s", show_choices=True)
        if choice == "n":
            new_name = Prompt.ask("  Search Comic Vine for series").strip()
            return ("search_series", new_name) if new_name else "research"
        if choice == "i":
            raw = Prompt.ask("  Comic Vine volume ID (numeric, from the URL)").strip()
            if raw.isdigit():
                return ("volume_id", int(raw))
            console.print("[yellow]  Not a valid ID — falling back to re-search.[/yellow]")
            return "research"
        return {"s": "n_nochange", "r": "research", "q": "q"}.get(choice, "n_nochange")

    if dry_run:
        console.print("[yellow](dry-run mode — no files will be changed)[/yellow]")
        return "y"

    console.print(
        "\n[dim]  y = apply   a = apply all remaining   n = skip"
        "   r = re-search (same name, bypass cache)"
        "   s = search new series name   i = enter Comic Vine volume ID   q = quit[/dim]"
    )
    choice = Prompt.ask(
        "[bold]Apply changes?[/bold]",
        choices=["y", "a", "n", "r", "s", "i", "q"],
        default="y",
        show_choices=True,
    )
    if choice == "s":
        new_name = Prompt.ask("  Search Comic Vine for series").strip()
        return ("search_series", new_name) if new_name else "research"
    if choice == "i":
        raw = Prompt.ask("  Comic Vine volume ID (numeric, from the URL)").strip()
        if raw.isdigit():
            return ("volume_id", int(raw))
        console.print("[yellow]  Not a valid ID — falling back to re-search.[/yellow]")
        return "research"
    return choice


# ---------------------------------------------------------------------------
# File renaming
# ---------------------------------------------------------------------------

def _safe_filename_part(text: str) -> str:
    """Sanitise a string for use as part of a filename.

    Path separators (/ and \\) are replaced with ' - ' so that crossover
    titles like "Star Trek/Green Lantern" become "Star Trek - Green Lantern"
    rather than being silently swallowed or, worse, creating a subdirectory.
    Other filesystem-unsafe characters are removed outright.
    """
    # Replace path separators with a readable dash separator
    text = re.sub(r"\s*/[/\\]+\s*", " - ", text)   # " / " or "/" → " - "
    text = re.sub(r"[/\\]", " - ", text)             # any remaining bare slash
    # Remove characters that are illegal on Windows/macOS/Linux
    text = re.sub(r'[*?:"<>|]', "", text)
    # Normalise whitespace and clean up stray dashes left by the replacements
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r" -\s*$", "", text)   # trailing " -" if slash was at end
    return text.strip()


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

    name = _safe_filename_part(series)
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


def _rename_if_needed(path: Path, metadata: dict, dry_run: bool) -> Path:
    """Rename file if its name doesn't already match the canonical form."""
    expected = canonical_filename(metadata, path)
    if path.name != expected:
        return rename_file(path, metadata, dry_run)
    return path


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
    smart_match: bool = False,
    full_metadata: bool = False,
    changelog: Optional[List] = None,
    volume_overrides: Optional[dict] = None,  # series_key → Comic Vine volume dict
) -> str:
    """
    Process a single comic file.
    Returns: 'updated', 'skipped', 'no_change', 'error', 'quit', 'converted'
    """
    is_cbr = path.suffix.lower() == ".cbr"
    apply_all = False  # set to True when user picks 'a' (apply all remaining)

    # Count sibling comic files — used to highlight matching series in the volume picker
    _comic_exts = {".cbz", ".cbr"}
    dir_comic_count: Optional[int] = sum(
        1 for f in path.parent.iterdir()
        if f.is_file() and f.suffix.lower() in _comic_exts
    )

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

    # --- Extract cover image for smart matching (only when enabled) ---
    cover_bytes: Optional[bytes] = extract_cover_image(path) if smart_match else None

    # --- Fetch from API (loop allows re-search / ID override) ---
    skip_cache = no_cache
    series_key = slugify_series(inferred.get("Series", "")).lower()
    # Wildcard "*" wins over series-specific key (set by --volume-id flag)
    overrides = volume_overrides or {}
    active_volume = overrides.get("*") or overrides.get(series_key)

    while True:
        proposed_meta, source = fetch_metadata(
            inferred, api_key, cache,
            skip_cache=skip_cache,
            volume_override=active_volume,
            cover_bytes=cover_bytes,
            full_metadata=full_metadata,
        )
        if proposed_meta is None:
            console.print(
                f"\n[yellow]Could not find metadata for:[/yellow] {path.name}\n"
                f"[dim]Inferred series: \"{inferred.get('Series', '?')}\"[/dim]"
            )
            if auto:
                console.print("[dim]Skipping (auto mode).[/dim]")
                return "skipped"

            # Interactive recovery — offer the same escape hatches as the confirmation UI
            console.print(
                "[dim]  s = skip   n = search Comic Vine for a different name"
                "   i = enter volume ID   q = quit[/dim]"
            )
            not_found_choice = Prompt.ask(
                "", choices=["s", "n", "i", "q"], default="s", show_choices=True
            )
            if not_found_choice == "q":
                return "quit"
            if not_found_choice == "s":
                return "skipped"
            if not_found_choice == "n":
                if not api_key:
                    console.print("[yellow]  No Comic Vine API key — cannot search.[/yellow]")
                    return "skipped"
                new_name = Prompt.ask("  Search Comic Vine for series").strip()
                if not new_name:
                    return "skipped"
                console.print(f"[dim]Searching Comic Vine for:[/dim] {new_name}")
                results = search_comicvine_volumes_all(new_name, api_key)
                if not results:
                    console.print(f"[yellow]  No results for '{new_name}'.[/yellow]")
                    return "skipped"
                new_volume = _pick_volume(results, highlight_count=dir_comic_count, inferred=inferred) if len(results) > 1 else results[0]
                pub = (new_volume.get("publisher") or {}).get("name", "?")
                console.print(
                    f"[green]Using:[/green] {new_volume.get('name')} "
                    f"({new_volume.get('start_year')}, {pub}, "
                    f"{new_volume.get('count_of_issues')} issues)"
                )
                active_volume = new_volume
                if volume_overrides is not None:
                    volume_overrides[series_key] = active_volume
                skip_cache = True
                continue   # retry fetch_metadata with the new active_volume
            if not_found_choice == "i":
                if not api_key:
                    console.print("[yellow]  No Comic Vine API key — cannot fetch by ID.[/yellow]")
                    return "skipped"
                raw_id = Prompt.ask("  Comic Vine volume ID (numeric, from the URL)").strip()
                if not raw_id.isdigit():
                    console.print("[yellow]  Not a valid ID.[/yellow]")
                    return "skipped"
                vol_id_input = int(raw_id)
                console.print(f"[dim]Fetching Comic Vine volume ID {vol_id_input}…[/dim]")
                new_volume = fetch_comicvine_volume_by_id(vol_id_input, api_key, cache)
                if not new_volume:
                    console.print(f"[yellow]  Could not fetch volume ID {vol_id_input}. Check the ID.[/yellow]")
                    return "skipped"
                pub = (new_volume.get("publisher") or {}).get("name", "?")
                console.print(
                    f"[green]Matched:[/green] {new_volume.get('name')} "
                    f"({new_volume.get('start_year')}, {pub}, "
                    f"{new_volume.get('count_of_issues')} issues)"
                )
                active_volume = new_volume
                if volume_overrides is not None:
                    volume_overrides[series_key] = active_volume
                skip_cache = True
                continue   # retry fetch_metadata with the new active_volume

        # Pop the CV page count before it can leak into the archive or the diff UI
        cv_page_count: Optional[int] = proposed_meta.pop("_cv_page_count", None)

        # Fill in page count from actual archive (always authoritative)
        page_count = get_page_count(path)
        if page_count:
            proposed_meta["PageCount"] = str(page_count)

        # Warn when the actual archive count differs meaningfully from CV's value.
        # A delta of ±1–4 is normal (ads, letters pages, covers); beyond that
        # it likely means pages are missing from the scan or the wrong issue matched.
        if cv_page_count and page_count:
            delta = page_count - cv_page_count
            if abs(delta) > 4:
                if delta < 0:
                    console.print(
                        f"[yellow]⚠  Page count mismatch:[/yellow] file has [bold]{page_count}[/bold] pages, "
                        f"CV reports [bold]{cv_page_count}[/bold] "
                        f"([red]delta {delta:+d}[/red]) — scan may be missing pages."
                    )
                else:
                    console.print(
                        f"[dim]ℹ  Page count: file has {page_count} pages, "
                        f"CV reports {cv_page_count} (delta {delta:+d}).[/dim]"
                    )

        # Preserve fields not returned by API that are already set,
        # except fields where a missing API value is meaningful.
        for field in METADATA_FIELDS:
            if field not in proposed_meta and field not in FIELDS_NO_PRESERVE and current_meta.get(field):
                proposed_meta[field] = current_meta[field]

        # --- Confirmation UI ---
        if auto:
            verdict, changes = show_confirmation_ui_auto(path, current_meta, proposed_meta, source)
            if verdict == "no_change":
                if not no_rename:
                    _rename_if_needed(path, proposed_meta, dry_run)
                return "no_change"
            if changelog is not None and changes:
                changelog.append((path, changes))
            break
        else:
            choice = show_confirmation_ui(path, current_meta, proposed_meta, source, dry_run)

            if choice == "q":
                return "quit"

            if choice == "research":
                skip_cache = True
                active_volume = None
                continue

            if isinstance(choice, tuple) and choice[0] == "search_series":
                # User typed a new series name — fetch all results and show a picker
                new_name = choice[1]
                console.print(f"[dim]Searching Comic Vine for:[/dim] {new_name}")
                results = search_comicvine_volumes_all(new_name, api_key)
                if not results:
                    console.print(f"[yellow]No results for '{new_name}'. Try a different name or enter an ID with [i].[/yellow]")
                    continue
                new_volume = _pick_volume(results, highlight_count=dir_comic_count, inferred=inferred) if len(results) > 1 else results[0]
                pub = (new_volume.get("publisher") or {}).get("name", "?")
                console.print(
                    f"[green]Using:[/green] {new_volume.get('name')} "
                    f"({new_volume.get('start_year')}, {pub}, "
                    f"{new_volume.get('count_of_issues')} issues)"
                )
                active_volume = new_volume
                # Save override so all subsequent files in this folder use the same volume
                if volume_overrides is not None:
                    volume_overrides[series_key] = active_volume
                skip_cache = True
                continue

            if isinstance(choice, tuple) and choice[0] == "volume_id":
                # User supplied a Comic Vine volume ID directly
                vol_id = choice[1]
                console.print(f"[dim]Fetching Comic Vine volume ID {vol_id}…[/dim]")
                new_volume = fetch_comicvine_volume_by_id(vol_id, api_key, cache)
                if not new_volume:
                    console.print(f"[yellow]Could not fetch volume ID {vol_id}. Check the ID and try again.[/yellow]")
                    continue
                pub = (new_volume.get("publisher") or {}).get("name", "?")
                console.print(
                    f"[green]Matched:[/green] {new_volume.get('name')} "
                    f"({new_volume.get('start_year')}, {pub}, "
                    f"{new_volume.get('count_of_issues')} issues)"
                )
                active_volume = new_volume
                if volume_overrides is not None:
                    volume_overrides[series_key] = active_volume
                skip_cache = True
                continue

            if choice == "n":
                return "skipped"
            if choice == "n_nochange":
                if not no_rename:
                    _rename_if_needed(path, proposed_meta, dry_run)
                return "no_change"
            if choice == "a":
                apply_all = True
            break  # "y" or "a" — proceed to apply

    # --- Apply ---
    ok = write_cbz_metadata(path, proposed_meta, dry_run)
    if not ok:
        return "error"

    # --- Rename ---
    if not no_rename:
        rename_file(path, proposed_meta, dry_run)

    return "apply_all" if apply_all else "updated"


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
# Mode: check-pages (page count integrity)
# ---------------------------------------------------------------------------

def _get_cv_page_count(meta: dict, cache: "DiskCache") -> Optional[int]:
    """
    Look up the Comic Vine page count for a comic from the local disk cache.
    Requires that comicrelief has already been run on the file (which populates
    the cache).  Returns None if no cached issue data is found.
    """
    series = meta.get("Series", "")
    number = meta.get("Number", "")
    year   = meta.get("Year", "")
    if not series or not number:
        return None

    search_name = slugify_series(series).lower()

    # Try to find a cached volume for this series (try with and without year)
    volume = None
    for yr in ([year, ""] if year else [""]):
        v = cache.get(f"cv_vol:{search_name}:{yr}")
        if v:
            volume = v
            break
        candidates = cache.get(f"cv_vol_list:{search_name}:{yr}")
        if candidates:
            volume = candidates[0]
            break

    if not volume:
        return None

    vol_id = volume.get("id")
    if not vol_id:
        return None

    # Normalise issue number to match the cache key written by fetch_comicvine_issue
    num_norm = (
        str(int(float(number)))
        if number.replace(".", "").isdigit()
        else number
    )
    issue = cache.get(f"cv_issue:{vol_id}:{num_norm}") or cache.get(f"cv_issue:{vol_id}:{number}")
    if not issue:
        return None

    pc = issue.get("page_count")
    try:
        val = int(pc)
        return val if val > 0 else None
    except (TypeError, ValueError):
        return None


def run_check_pages_mode(comics: List[Path], cache: "DiskCache") -> None:
    """
    Check actual image counts in CBZ/CBR files against:
      - Stored  : the PageCount field in ComicInfo.xml
      - CV      : Comic Vine's page_count (read from local cache — run
                  comicrelief on the folder first to populate it)

    Delta = Actual − CV.  A negative delta means pages may be missing.
    Note: CV counts include ads and letters pages; scans that skip ads
    will show a small negative delta even when complete.
    """
    table = Table(
        box=box.ROUNDED,
        show_header=True,
        header_style="bold",
        row_styles=["", "dim"],
        padding=(0, 1),
    )
    table.add_column("",       width=1,  no_wrap=True)
    table.add_column("File",   min_width=24, max_width=52, no_wrap=True)
    table.add_column("Fmt",    width=3,  no_wrap=True)
    table.add_column("Actual", width=6,  justify="right", no_wrap=True)
    table.add_column("Stored", width=6,  justify="right", no_wrap=True)
    table.add_column("CV",     width=6,  justify="right", no_wrap=True)
    table.add_column("Delta",  width=7,  justify="right", no_wrap=True)

    ok_count = mismatch_count = unknown_count = 0
    has_cv_data = False

    for path in comics:
        fmt    = path.suffix.upper().lstrip(".")
        actual = get_page_count(path)
        meta   = read_metadata(path)

        stored_str = meta.get("PageCount", "")
        stored: Optional[int] = int(stored_str) if stored_str and stored_str.isdigit() else None

        cv_count = _get_cv_page_count(meta, cache)
        if cv_count is not None:
            has_cv_data = True

        # Status — primary reference is CV; fall back to Stored if no CV data
        if cv_count is not None:
            is_ok: Optional[bool] = (actual == cv_count)
        elif stored is not None:
            is_ok = (actual == stored)
        else:
            is_ok = None  # no reference at all

        if is_ok is None:
            status = "[yellow]?[/yellow]"
            unknown_count += 1
        elif is_ok:
            status = "[green]✓[/green]"
            ok_count += 1
        else:
            status = "[red]✗[/red]"
            mismatch_count += 1

        actual_cell = str(actual) if actual else "[dim]—[/dim]"
        stored_cell = str(stored) if stored is not None else "[dim]—[/dim]"
        cv_cell     = str(cv_count) if cv_count is not None else "[dim]—[/dim]"

        if cv_count is not None:
            delta = actual - cv_count
            if delta == 0:
                delta_cell = "[green]0[/green]"
            elif delta < 0:
                delta_cell = f"[red]{delta}[/red]"
            else:
                delta_cell = f"[yellow]+{delta}[/yellow]"
        elif stored is not None and actual != stored:
            delta = actual - stored
            delta_cell = f"[red]{delta}[/red]" if delta < 0 else f"[yellow]+{delta}[/yellow]"
        else:
            delta_cell = "[dim]—[/dim]"

        fname = path.name
        if len(fname) > 52:
            fname = fname[:51] + "…"

        table.add_row(status, fname, fmt, actual_cell, stored_cell, cv_cell, delta_cell)

    console.print(table)
    console.print(
        f"\n{len(comics)} file(s)   "
        f"[green]✓ {ok_count} OK[/green]   "
        f"[red]✗ {mismatch_count} mismatch[/red]   "
        f"[yellow]? {unknown_count} no reference[/yellow]"
    )

    if not has_cv_data:
        console.print(
            "\n[dim]CV column is empty — run comicrelief on this folder first "
            "to populate the Comic Vine page-count cache.[/dim]"
        )
    elif mismatch_count:
        console.print(
            "\n[dim]Tip: CV counts include ads and letters pages. "
            "A small negative delta (−1 to −5) may be expected if your scans skip ads.[/dim]"
        )


# ---------------------------------------------------------------------------
# Mode: list (display metadata table)
# ---------------------------------------------------------------------------

# Symbols for quick visual health check
_TICK  = "[green]✓[/green]"
_CROSS = "[red]✗[/red]"
# Default core fields — presence of all of these earns a ✓
DEFAULT_CORE_FIELDS = {"Series", "Number", "Year", "Publisher"}


def _dominant_volume(series_metas: List[dict]) -> Optional[str]:
    """
    Return the most common non-empty Volume value across a series.
    Used to detect files where Volume was incorrectly set to equal the issue number.
    """
    from collections import Counter
    counts: Counter = Counter()
    for meta in series_metas:
        v = meta.get("Volume", "").strip()
        if v:
            counts[v] += 1
    if not counts:
        return None
    return counts.most_common(1)[0][0]


def run_list_mode(
    comics: List[Path],
    core_fields: Optional[set] = None,
    display_fields: Optional[List[str]] = None,
) -> None:
    """Display a per-file metadata table followed by a per-series collection summary.

    display_fields controls which metadata columns appear (order matters).
    If None, defaults to DEFAULT_LIST_FIELDS plus any core_fields not already included.
    """
    if core_fields is None:
        core_fields = DEFAULT_CORE_FIELDS

    # Build the column list:
    # - If caller explicitly passed display_fields, use exactly that.
    # - Otherwise use defaults and auto-append any core fields not already present.
    if display_fields is not None:
        col_fields = list(display_fields)
    else:
        col_fields = list(DEFAULT_LIST_FIELDS)
        for cf in sorted(core_fields):
            if cf not in col_fields:
                col_fields.append(cf)

    # --- Read all metadata up front so we only hit each file once ---
    all_meta: List[Tuple[Path, dict]] = []
    for path in comics:
        meta = read_metadata(path)
        if not meta.get("Series") or not meta.get("Number"):
            inferred = parse_filename(path)
            for k, v in inferred.items():
                if not meta.get(k):
                    meta[k] = v
        all_meta.append((path, meta))

    # --- Pre-compute per-series dominant Volume for inconsistency detection ---
    from collections import defaultdict as _dd
    series_to_metas: dict = _dd(list)
    for _, meta in all_meta:
        key = meta.get("Series", "").lower() or "__unknown__"
        series_to_metas[key].append(meta)
    dominant_vol: dict = {k: _dominant_volume(v) for k, v in series_to_metas.items()}

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
    # Fixed columns always present
    file_table.add_column("",     width=1,  no_wrap=True)
    file_table.add_column("File", min_width=20, max_width=40, no_wrap=True)
    file_table.add_column("Fmt",  width=3,  no_wrap=True)

    # Dynamic metadata columns
    for field in col_fields:
        spec = FIELD_COLUMN_SPECS.get(field)
        if spec:
            header, kwargs, _ = spec
            file_table.add_column(header, **kwargs)
        else:
            # Unknown/unlisted field — generic column
            file_table.add_column(field, min_width=8, max_width=20, no_wrap=True)

    def cell(key: str, max_len: Optional[int], _meta: dict) -> str:
        val = _meta.get(key, "")
        if not val:
            return "[dim]—[/dim]"
        if max_len is None:
            return val
        return val[:max_len - 1] + "…" if len(val) >= max_len else val

    missing_core = 0
    for path, meta in all_meta:
        fmt = path.suffix.upper().lstrip(".")
        has_all_core = all(meta.get(f) for f in core_fields)
        indicator = _TICK if has_all_core else _CROSS
        if not has_all_core:
            missing_core += 1

        # Pre-compute Vol suspicious flag (needed whether or not Volume is in col_fields)
        vol_val = meta.get("Volume", "").strip()
        num_val = meta.get("Number", "").strip()
        s_key   = meta.get("Series", "").lower() or "__unknown__"
        dom_vol = dominant_vol.get(s_key)
        vol_suspicious = bool(
            vol_val and num_val and vol_val == num_val and num_val != "1"
        ) or bool(
            vol_val and dom_vol and vol_val != dom_vol
        ) or bool(
            not vol_val and dom_vol  # missing when others have it
        )

        row_cells = [
            indicator,
            path.name[:39] + "…" if len(path.name) > 40 else path.name,
            fmt,
        ]

        for field in col_fields:
            if field == "Volume":
                # Special yellow highlight for suspicious Vol values
                if vol_suspicious:
                    row_cells.append(f"[yellow]{vol_val}[/yellow]" if vol_val else "[yellow]—[/yellow]")
                else:
                    spec = FIELD_COLUMN_SPECS.get("Volume")
                    row_cells.append(cell("Volume", spec[2] if spec else 4, meta))
            elif field == "Year":
                # Year: never truncate
                row_cells.append(meta.get("Year", "") or "[dim]—[/dim]")
            else:
                spec = FIELD_COLUMN_SPECS.get(field)
                max_len = spec[2] if spec else 20
                row_cells.append(cell(field, max_len, meta))

        file_table.add_row(*row_cells)

    core_label = " / ".join(sorted(core_fields))
    console.print(file_table)
    console.print(
        f"\n{len(comics)} file(s)  "
        f"[green]✓ {len(comics) - missing_core} complete[/green]   "
        f"[red]✗ {missing_core} missing core fields[/red] "
        f"[dim]({core_label})[/dim]"
    )
    if any(
        (meta.get("Volume", "") and meta.get("Number", "") and meta.get("Volume") == meta.get("Number") and meta.get("Number") != "1")
        or (meta.get("Volume", "") != (dominant_vol.get((meta.get("Series","").lower() or "__unknown__")) or meta.get("Volume","")))
        for _, meta in all_meta
    ):
        console.print("  [yellow]⚠ Yellow Vol cells indicate inconsistent or suspicious volume numbers.[/yellow]")

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
    mode_group.add_argument(
        "--check-pages", action="store_true",
        help=(
            "Check actual image counts against stored and Comic Vine page counts. "
            "Run comicrelief on the folder first to populate the CV cache."
        ),
    )

    # General options
    parser.add_argument("--dry-run", action="store_true", help="Show changes without writing")
    parser.add_argument("--no-rename", action="store_true", help="Do not rename files")
    parser.add_argument("--no-cache", action="store_true", help="Ignore cached API results and re-fetch")
    parser.add_argument("--smart-match", action="store_true", help="Use cover image comparison to disambiguate series with similar names (requires Pillow and imagehash)")
    parser.add_argument("--full-metadata", action="store_true", help="Fetch and store all available metadata fields: Genre, Tags (locations/teams), LanguageISO, all characters (no cap), and a summary fallback from the volume description")
    parser.add_argument("--api-key", help="Comic Vine API key")
    parser.add_argument("--cache-file", default=str(DEFAULT_CACHE_PATH), help="Path to API cache JSON file")
    parser.add_argument(
        "--core-fields",
        default=None,
        help=(
            "Comma-separated fields that must be present for a ✓ in --list mode. "
            f"Default: {','.join(sorted(DEFAULT_CORE_FIELDS))}. "
            f"Available: {', '.join(METADATA_FIELDS)}"
        ),
    )
    parser.add_argument(
        "--fields",
        default=None,
        metavar="FIELD[,FIELD…]",
        help=(
            "Comma-separated list of metadata columns to show in --list mode, "
            "overriding the defaults. "
            f"Default columns: {','.join(DEFAULT_LIST_FIELDS)}. "
            f"Available: {', '.join(METADATA_FIELDS)}"
        ),
    )
    parser.add_argument(
        "--volume-id",
        type=int,
        default=None,
        metavar="ID",
        help=(
            "Force a specific Comic Vine volume ID for every file in this run. "
            "Find the ID in the URL: comicvine.gamespot.com/…/4050-ID/. "
            "Useful when auto-matching picks the wrong series."
        ),
    )
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
        if args.core_fields:
            raw = [f.strip() for f in args.core_fields.split(",")]
            invalid = [f for f in raw if f not in METADATA_FIELDS]
            if invalid:
                console.print(f"[red]Unknown field(s) in --core-fields:[/red] {', '.join(invalid)}")
                console.print(f"[dim]Available: {', '.join(METADATA_FIELDS)}[/dim]")
                sys.exit(1)
            core_fields = set(raw)
        else:
            core_fields = None

        display_fields: Optional[List[str]] = None
        if args.fields:
            raw_fields = [f.strip() for f in args.fields.split(",")]
            invalid_fields = [f for f in raw_fields if f not in METADATA_FIELDS]
            if invalid_fields:
                console.print(f"[red]Unknown field(s) in --fields:[/red] {', '.join(invalid_fields)}")
                console.print(f"[dim]Available: {', '.join(METADATA_FIELDS)}[/dim]")
                sys.exit(1)
            display_fields = raw_fields

        run_list_mode(comics, core_fields=core_fields, display_fields=display_fields)
        return

    # -------------------------------------------------------------------------
    # Mode: --check-pages
    # -------------------------------------------------------------------------
    if args.check_pages:
        cache = DiskCache(Path(args.cache_file))
        run_check_pages_mode(comics, cache)
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
    volume_overrides: dict = {}  # series_key → CV volume dict; sticky across files in a run

    # --volume-id: resolve the volume once and use it for every file in this run
    if args.volume_id:
        if not api_key:
            console.print("[red]--volume-id requires a Comic Vine API key.[/red]")
            sys.exit(1)
        console.print(f"[dim]Fetching Comic Vine volume ID {args.volume_id}…[/dim]")
        forced_volume = fetch_comicvine_volume_by_id(args.volume_id, api_key, cache)
        if not forced_volume:
            console.print(f"[red]Could not fetch Comic Vine volume ID {args.volume_id}. Check the ID.[/red]")
            sys.exit(1)
        pub = (forced_volume.get("publisher") or {}).get("name", "?")
        console.print(
            f"[green]Volume override:[/green] {forced_volume.get('name')} "
            f"({forced_volume.get('start_year')}, {pub}, {forced_volume.get('count_of_issues')} issues)\n"
        )
        # Store under the wildcard key "*" — process_file checks this before the series key
        volume_overrides["*"] = forced_volume

    force_auto = False  # becomes True when user picks 'a' mid-review

    for i, comic_path in enumerate(comics, 1):
        is_auto = args.auto or force_auto
        if is_auto:
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
                auto=is_auto,
                smart_match=args.smart_match,
                full_metadata=args.full_metadata,
                changelog=changelog,
                volume_overrides=volume_overrides,
            )
        except KeyboardInterrupt:
            console.print("\n[yellow]Interrupted by user.[/yellow]")
            break
        except Exception as e:
            console.print(f"[red]Unexpected error processing {comic_path.name}: {e}[/red]")
            result = "error"

        if result in ("updated", "apply_all"):
            stats["updated"] += 1
            if result == "apply_all":
                force_auto = True
                remaining = len(comics) - i
                console.print(
                    f"\n[cyan]Switching to auto mode — "
                    f"applying {remaining} remaining file(s) without confirmation.[/cyan]\n"
                )
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
    if (args.auto or force_auto) and changelog:
        console.rule()
        console.print(f"\n[bold]Changes applied ({len(changelog)} file(s)):[/bold]")
        print_auto_changelog(changelog)

    print_summary(stats, error_files, ambiguous_files)


if __name__ == "__main__":
    main()
