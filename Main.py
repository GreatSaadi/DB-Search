"""
╔══════════════════════════════════════════════════════════════╗
║        DATABASE SEARCHER  v2.0.0  —  @GreatSaadi             ║
║              https://github.com/GreatSaadi                   ║
║                       DB Searcher                            ║
╚══════════════════════════════════════════════════════════════╝
"""

import os, sys, sqlite3, hashlib, json, time, shutil, threading, unicodedata, re, csv
from functools import lru_cache as _lru_cache

# ══════════════════════════════════════════════════════════════
#  AUTO-DEPENDENCY BOOTSTRAP
#  Checks every third-party package this script needs and, if any
#  is missing, installs it automatically with pip (no manual steps
#  for the end user — this is what lets a plain double-click / exe
#  run work on a machine that never had these libraries). After a
#  successful install we re-exec the same interpreter so the newly
#  installed packages are importable in a fresh process (pip
#  installs are not always visible to an already-running one).
# ══════════════════════════════════════════════════════════════
_REQUIRED_PACKAGES = {
    "pandas":         "pandas",
    "rich":           "rich",
    "questionary":    "questionary",
    "openpyxl":       "openpyxl",       # needed by pandas to read .xlsx
    "arabic_reshaper": "arabic-reshaper",
    "bidi":           "python-bidi",
}

def _bootstrap_dependencies():
    import importlib
    missing_pip_names = []
    for import_name, pip_name in _REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing_pip_names.append(pip_name)

    if not missing_pip_names:
        return

    print(f"[Setup] Missing packages detected: {', '.join(missing_pip_names)}")
    print("[Setup] Installing automatically, please wait...")
    try:
        import subprocess
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "--quiet",
            "--disable-pip-version-check", *missing_pip_names,
        ])
    except Exception as e:
        print(f"[Setup] Automatic install failed: {e}")
        print(f"[Setup] Please run manually:  pip install {' '.join(missing_pip_names)}")
        sys.exit(1)

    print("[Setup] Done — restarting...")
    try:
        os.execv(sys.executable, [sys.executable] + sys.argv)
    except Exception:
        # os.execv can fail in some frozen/exe contexts; fall back to a
        # plain process relaunch instead of leaving the user stuck.
        import subprocess
        subprocess.check_call([sys.executable] + sys.argv)
        sys.exit(0)

if not getattr(sys, "frozen", False):
    # Skip the bootstrap entirely when running as a PyInstaller-built
    # .exe: a frozen build already bundles every dependency inside
    # itself, so there is nothing to pip-install and no Python
    # environment on the end user's machine to install it into.
    _bootstrap_dependencies()

import pandas as pd
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, wait, FIRST_COMPLETED
from rich.console   import Console
from rich.table     import Table
from rich.panel     import Panel
from rich.align     import Align
from rich.progress  import (Progress, SpinnerColumn, TextColumn,
                             BarColumn, MofNCompleteColumn, TimeElapsedColumn,
                             TaskProgressColumn)
from rich.rule      import Rule
from rich           import box
import questionary

# ── RTL / Persian display fix ─────────────────────────────────
try:
    import arabic_reshaper
    from bidi.algorithm import get_display as _bidi_display
    _RTL_CHAR_RE = re.compile(r"[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]")
    def fix_rtl(text: str) -> str:
        if not text or not isinstance(text, str): return text
        if not _RTL_CHAR_RE.search(text):
            return text   # pure Latin/numeric — skip reshaping, cheaper + nothing to fix
        # FIX: base_dir="R" is required here. get_display() auto-detects
        # direction from the first *strong* character in the string; a cell
        # like "09123456789 - علی رضایی" starts with digits (neutral/weak),
        # so autodetect can pick the wrong base direction and reorder the
        # Persian segment incorrectly — this is what caused the scrambled
        # look even with the libraries installed. Forcing "R" whenever the
        # cell contains any RTL character fixes that.
        try:
            return _bidi_display(arabic_reshaper.reshape(text), base_dir="R")
        except TypeError:
            # older python-bidi versions may not accept base_dir as kwarg here
            return _bidi_display(arabic_reshaper.reshape(text))
    _HAS_BIDI = True
except ImportError:
    _HAS_BIDI = False
    def fix_rtl(text: str) -> str: return text

# ── Windows UTF-8 setup ───────────────────────────────────────
if os.name == "nt":
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
    except Exception:
        pass
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

console = Console(highlight=False, force_terminal=True, markup=True)

# ══════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "cli_settings.json")
ENCODINGS   = ["utf-8-sig", "utf-8", "cp1256", "windows-1256", "latin-1"]
_CPU        = os.cpu_count() or 4

ALL_EXTS   = {".csv", ".txt", ".xlsx", ".json", ".db", ".sqlite"}
_FAST_EXTS = {".csv", ".txt"}
_DELIM_CACHE = {}

# ── Text normalisation ────────────────────────────────────────
_CHAR_MAP = {ord(k): v for k, v in {
    "ي": "ی", "ى": "ی", "ئ": "ی", "ك": "ک",
    "ة": "ه", "أ": "ا", "إ": "ا", "آ": "ا", "ؤ": "و"
}.items()}
_DIACRITICS_RE = re.compile(r"[\u064B-\u065F\u0670\u06D6-\u06ED]")

def normalize_text(s: str) -> str:
    if not isinstance(s, str): s = str(s)
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_CHAR_MAP)
    s = _DIACRITICS_RE.sub("", s)
    return s.strip().lower()

# ══════════════════════════════════════════════════════════════
#  COLOURS
# ══════════════════════════════════════════════════════════════
C_PINK    = "#ff79c6"
C_CYAN    = "#8be9fd"
C_CYAN_LT = "#c8f6ff"
C_CYAN_MD = "#62e0ff"
C_CYAN_DK = "#22b8cf"
C_GREEN   = "#50fa7b"
C_LIME    = "#69ff94"
C_YELLOW  = "#f1fa8c"
C_GOLD    = "#ffd700"
C_PURPLE  = "#bd93f9"
C_ORANGE  = "#ffb86c"
C_RED     = "#ff5555"
C_DIM     = "#6272a4"
C_WHITE   = "#f8f8f2"
C_TEAL    = "#1de9b6"
C_MAGENTA = "#ff00ff"

# ══════════════════════════════════════════════════════════════
#  UI KIT  —  small reusable print helpers (keeps message styling
#  consistent everywhere instead of re-typing colour tags at every
#  call site)
# ══════════════════════════════════════════════════════════════
def _ok(msg: str, pause: float = 0.0):
    console.print(f"\n  [bold {C_LIME}]✔[/bold {C_LIME}]  [{C_WHITE}]{msg}[/{C_WHITE}]")
    if pause: time.sleep(pause)

def _warn(msg: str, pause: float = 0.0):
    console.print(f"\n  [bold {C_YELLOW}]⚠[/bold {C_YELLOW}]  [{C_YELLOW}]{msg}[/{C_YELLOW}]")
    if pause: time.sleep(pause)

def _err(msg: str, pause: float = 0.0):
    console.print(f"\n  [bold {C_RED}]✕[/bold {C_RED}]  [{C_RED}]{msg}[/{C_RED}]")
    if pause: time.sleep(pause)

def _dim(msg: str):
    console.print(f"  [{C_DIM}]{msg}[/{C_DIM}]")

def _gradient_rule(width: int = 70, ornament: str = "◆"):
    """A symmetric divider that brightens from both edges toward a gold
    ornament in the middle — like a little sunrise under the logo — instead
    of one flat-coloured dash. Defaults to the logo's own width (70) so it
    never runs short/long against it regardless of terminal size."""
    ramp = ["#0f6b80", "#157f99", "#1a94b3", "#22b8cf", "#3ecbe0",
            "#62e0ff", "#8be9fd", "#c8f6ff"]
    half = max(1, (width - 2) // 2)
    each = max(1, half // len(ramp))

    def _side(colours):
        return "".join(f"[{c}]" + "─" * each + f"[/{c}]" for c in colours)

    left  = _side(ramp)
    right = _side(list(reversed(ramp)))
    return f"  {left} [bold {C_GOLD}]{ornament}[/bold {C_GOLD}] {right}"

def _speed_badge(mb_s: float) -> str:
    if mb_s >= 400:  return f"[bold {C_MAGENTA}]🚀 BLAZING[/bold {C_MAGENTA}]"
    if mb_s >= 150:  return f"[bold {C_LIME}]⚡ FAST[/bold {C_LIME}]"
    if mb_s >= 40:   return f"[bold {C_CYAN_MD}]🌊 STEADY[/bold {C_CYAN_MD}]"
    return f"[bold {C_ORANGE}]🐢 STEADY-SLOW[/bold {C_ORANGE}]"

# ══════════════════════════════════════════════════════════════
#  CONFIG LOAD / SAVE
# ══════════════════════════════════════════════════════════════
def load_config():
    d = {
        "target_dir":  "./",
        "use_cache":   False,
        "cache_dir":   os.path.join(SCRIPT_DIR, "db_searcher_cache"),
        "max_workers": None,
        "active_exts": [".csv"],
        "chunk_mb":    64,
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                saved = json.load(f)
                d.update(saved)
                if not isinstance(d["active_exts"], list):
                    d["active_exts"] = [".csv"]
        except Exception:
            pass
    return d

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

config = load_config()

def _workers() -> int:
    w = config.get("max_workers")
    return _CPU if not w else max(1, int(w))

def _active_exts() -> set:
    return set(config.get("active_exts", [".csv"]))

# ══════════════════════════════════════════════════════════════
#  QUESTIONARY STYLE
# ══════════════════════════════════════════════════════════════
_STYLE = questionary.Style([
    ("qmark",       f"fg:{C_TEAL} bold"),
    ("question",    f"fg:{C_CYAN_LT} bold"),
    ("answer",      f"fg:{C_LIME} bold"),
    ("pointer",     f"fg:{C_PINK} bold"),
    ("highlighted", f"fg:{C_GOLD} bold"),
    ("selected",    f"fg:{C_LIME}"),
    ("instruction", f"fg:{C_DIM}"),
    ("text",        f"fg:{C_CYAN_MD}"),
    ("separator",   f"fg:{C_DIM}"),
])

# ══════════════════════════════════════════════════════════════
#  LOGO  —  gradient pink→purple "DB SEARCH"
# ══════════════════════════════════════════════════════════════
_LOGO_ROWS = [
    ("██████╗ ██████╗      ██████╗ ███████╗ █████╗ ██████╗   ██████╗██╗  ██╗", "#9bfdff"),
    ("██╔══██╗██╔══██╗     ██╔═══╝ ██╔════╝██╔══██╗██╔══██╗ ██╔════╝██║  ██║", "#7cf6ff"),
    ("██║  ██║██████╔╝     ╚█████╗ █████╗  ███████║██████╔╝ ██║     ███████║", "#62e0ff"),
    ("██║  ██║██╔══██╗      ╚═══██╗██╔══╝  ██╔══██║██╔══██╗ ██║     ██╔══██║", "#48d4ff"),
    ("██████╔╝██████╔╝     ██████╔╝███████╗██║  ██║██║  ██║ ╚██████╗██║  ██║", "#33c5fa"),
    ("╚═════╝ ╚═════╝      ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝  ╚═════╝╚═╝  ╚═╝", "#22b8cf"),
]
def print_header():
    os.system("cls" if os.name == "nt" else "clear")   # FIX: OS-conditional clear (cls broke on Linux/macOS, console.clear() unreliable on some terminals)
    console.print()
    for line, colour in _LOGO_ROWS:
        console.print(f"  [bold {colour}]{line}[/bold {colour}]")
    console.print()
    # FIX: was Align.center() — on wide terminals that pushes the byline
    # far right while the gradient rule below stays left-anchored under the
    # logo, so the divider ends up looking short/cut-off relative to it.
    # Left-aligned "chips" under the logo's own indent fixes that and reads
    # more like a real app's status strip than plain centred text.
    console.print(
        f"  [bold {C_TEAL}]⟡[/bold {C_TEAL}] [bold {C_CYAN}]v1.0[/bold {C_CYAN}]"
        f"   [{C_DIM}]│[/{C_DIM}]   [bold {C_TEAL}]⟡[/bold {C_TEAL}]"
        f" [italic bold {C_CYAN_LT}]@GreatSaadi[/italic bold {C_CYAN_LT}]"
        f"   [{C_DIM}]│[/{C_DIM}]   [bold {C_TEAL}]⟡[/bold {C_TEAL}]"
        f" [{C_CYAN_LT}]github.com/GreatSaadi[/{C_CYAN_LT}]"
    )
    console.print(
        f"  [italic {C_DIM}]⚡ Blazing multi-format search"
        f"  ·  CSV · TXT · XLSX · JSON · SQLite[/italic {C_DIM}]"
    )
    console.print(_gradient_rule())
    console.print(_status_panel())
    console.print()
# ── helpers ───────────────────────────────────────────────────
def _cache_stats():
    cdir = os.path.abspath(config["cache_dir"])
    n, sz = 0, 0
    if os.path.isdir(cdir):
        for e in os.scandir(cdir):
            if e.name.endswith(".pkl"):
                n += 1; sz += e.stat().st_size
    label = (f"{sz/1_048_576:.1f} MB" if sz >= 1_048_576
             else f"{sz//1024} KB" if sz >= 1024 else f"{sz} B")
    return n, label

def _ext_badge() -> str:
    exts = sorted(_active_exts())
    if set(exts) <= _FAST_EXTS:
        return (f"[bold {C_LIME}]⚡ FAST  "
                f"({', '.join(e.lstrip('.').upper() for e in exts)})[/bold {C_LIME}]")
    label = "  ".join(e.lstrip(".").upper() for e in exts)
    return f"[bold {C_CYAN_MD}]{label}[/bold {C_CYAN_MD}]"

def _thread_bar(fill: int, total: int) -> str:
    return (f"[bold {C_TEAL}]" + "█" * fill + f"[/bold {C_TEAL}]"
          + f"[{C_DIM}]"       + "░" * max(0, total - fill) + f"[/{C_DIM}]")

# ── status panel (header) — NO configuration duplicate ────────
def _status_panel() -> Panel:
    c_on = config["use_cache"]
    w    = config.get("max_workers")
    tgt  = os.path.abspath(config["target_dir"])
    cdir = os.path.abspath(config["cache_dir"])
    n_files, csize = _cache_stats()

    w_fill  = _CPU if not w else min(int(w), _CPU)
    w_label = (f"[bold {C_LIME}]ALL {_CPU} CORES[/bold {C_LIME}]" if not w
               else f"[bold {C_CYAN_MD}]{w} threads[/bold {C_CYAN_MD}]")
    tbar    = _thread_bar(w_fill, _CPU)

    cache_val = (
        f"[bold {C_LIME}]● ON[/bold {C_LIME}]  [{C_DIM}]{n_files} file(s) · {csize}[/{C_DIM}]"
        if c_on else
        f"[bold {C_YELLOW}]○ OFF[/bold {C_YELLOW}]  [{C_DIM}]live read[/{C_DIM}]"
    )

    grid = Table.grid(padding=(0, 1))
    grid.add_column(justify="left", style=f"bold {C_CYAN_DK}", no_wrap=True)
    grid.add_column(justify="left")
    grid.add_row("🎯 TARGET ", f"[{C_WHITE}]{tgt}[/{C_WHITE}]")
    grid.add_row("💾 CACHE  ", cache_val)
    grid.add_row("📂 CDIR   ", f"[{C_DIM}]{cdir}[/{C_DIM}]")
    grid.add_row("🧵 THREADS", f"{tbar}  {w_label}  [{C_DIM}]({_CPU} cores)[/{C_DIM}]")
    grid.add_row("📑 EXTS   ", _ext_badge())
    grid.add_row("🧩 CHUNK  ",
                 f"[{C_WHITE}]{config.get('chunk_mb', 64)} MB[/{C_WHITE}]"
                 f"  [{C_DIM}](fast-engine parallel split)[/{C_DIM}]")

    return Panel(grid,
                 title=f"[bold {C_CYAN_LT}]◈  SYSTEM STATUS[/bold {C_CYAN_LT}]",
                 border_style=C_CYAN_DK, box=box.ROUNDED, padding=(0, 1))

# ══════════════════════════════════════════════════════════════
#  FILE DISCOVERY
# ══════════════════════════════════════════════════════════════
def get_files(path):
    exts = _active_exts()
    out = []
    for root, _dirs, names in os.walk(path):
        for name in names:
            if name.startswith("~$") or name.startswith("."): continue
            if os.path.splitext(name)[1].lower() in exts:
                out.append(os.path.join(root, name))
    return sorted(set(out))

# ══════════════════════════════════════════════════════════════
#  FAST CSV/TXT ENGINE
# ══════════════════════════════════════════════════════════════

_ENC_CACHE = {}

def _detect_encoding(filepath: str) -> str:
    """Sniffs the real encoding ONCE per file with a strict decode test on a
    sample, instead of always opening with errors='ignore' (which silently
    swallows decode errors and effectively always "succeeds" with the first
    encoding in the list, even when it's wrong — the old code's ENCODINGS
    loop almost never actually fell through to cp1256/etc. because of this).
    Getting the real encoding right also lets us pre-filter on raw BYTES
    (see below) instead of decoding every line."""
    ap = os.path.abspath(filepath)
    cached = _ENC_CACHE.get(ap)
    if cached:
        return cached
    try:
        with open(filepath, "rb") as fh:
            sample = fh.read(262144)
    except Exception:
        _ENC_CACHE[ap] = "utf-8"
        return "utf-8"
    for enc in ENCODINGS:
        try:
            sample.decode(enc)
            _ENC_CACHE[ap] = enc
            return enc
        except (UnicodeDecodeError, LookupError):
            continue
    _ENC_CACHE[ap] = "utf-8"
    return "utf-8"


def _quick_normalize(s: str) -> str:
    """Cheap first-pass normaliser: folds Arabic/Persian character variants
    (ي→ی, ك→ک, …) and lowercases — WITHOUT the costlier NFKC pass or
    diacritic stripping that normalize_text() does. Used only as a fallback
    when the query has too many character-fold combinations for
    _query_variants() to enumerate (see below)."""
    if not isinstance(s, str): s = str(s)
    return s.translate(_CHAR_MAP).lower()

# Reverse of _CHAR_MAP: canonical char -> every raw form that folds to it.
_CHAR_VARIANTS = {}
for _k, _v in {
    "ي": "ی", "ى": "ی", "ئ": "ی", "ك": "ک",
    "ة": "ه", "أ": "ا", "إ": "ا", "آ": "ا", "ؤ": "و"
}.items():
    _CHAR_VARIANTS.setdefault(_v, {_v}).add(_k)
_MAX_QUERY_VARIANTS = 64

@_lru_cache(maxsize=512)
def _query_variants(term_norm: str):
    """Expands a normalized Persian/Arabic query into every raw spelling it
    could appear as in the file (ي vs ی, ك vs ک, ...), e.g. "علی" ->
    {"علی", "علي", "على", "علئ"}. This is the key speed fix: instead of
    normalizing every single line of a file (slow — a Python-level
    translate+lower call per line, on top of the file-read cost), we
    normalize the query ONCE and let the per-line check be a handful of
    plain C-level substring tests, which is what makes the fast engine
    actually fast for Persian. Returns None if the query has too many
    variant combinations (caller falls back to per-line normalization)."""
    options = [_CHAR_VARIANTS.get(ch, {ch}) for ch in term_norm]
    total = 1
    for o in options:
        total *= len(o)
        if total > _MAX_QUERY_VARIANTS:
            return None
    variants = [""]
    for o in options:
        variants = [v + c for v in variants for c in o]
    return variants


def _body_codec(enc: str) -> str:
    """'utf-8-sig' only strips/adds a BOM at the very start of a file — used
    per-line (for every data line, or when re-encoding the search term) it
    would wrongly prepend/strip 3 bytes on every single comparison, which
    silently breaks all byte-level matching. The header line (read once,
    also the only place a BOM could appear) is decoded separately with the
    real codec; every line after that uses the plain base codec."""
    return enc[:-4] if enc.endswith("-sig") else enc


_PATTERN_CACHE = {}

def _variant_byte_pattern(variants, enc: str):
    """Compiled BYTES regex that matches any raw spelling variant, encoded
    in the file's real codec. Searching raw bytes lets us reject almost
    every non-matching line WITHOUT decoding it — decoding gigabytes of
    text (UTF-8 validation, cp1256 table lookups, …) is the single biggest
    cost when scanning large files, and it was happening on EVERY line
    before, even ones that could never match.

    FIX (speed): a large file is split into many chunks (_split_offsets),
    and every chunk is a separate task that used to recompile this exact
    same regex from scratch — for a 15 GB file split into hundreds of
    64 MB chunks, that's hundreds of redundant re.compile() calls for the
    IDENTICAL pattern every single search. Worker processes stay alive
    and handle many chunks over their lifetime (ProcessPoolExecutor
    reuses them), so caching per (variants, enc) here means each distinct
    pattern is compiled once per worker process instead of once per
    chunk — the bigger the file, the bigger this win."""
    key = (tuple(variants), enc)
    if key in _PATTERN_CACHE:
        return _PATTERN_CACHE[key]
    parts = [re.escape(v.encode(enc, errors="ignore")) for v in variants if v]
    parts = [p for p in parts if p]
    compiled = re.compile(b"|".join(parts)) if parts else None
    _PATTERN_CACHE[key] = compiled
    return compiled


def _split_offsets(filepath: str, chunk_size: int):
    """Splits a file into line-aligned (start, end) byte ranges so it can be
    scanned by several worker PROCESSES at once. This is the other big
    speed fix: the old code submitted exactly one task per FILE, so with
    9 files you could never use more than 9 cores no matter how many were
    available, and a couple of oversized files would dominate the whole
    run while the rest of the cores sat idle. Splitting every file into
    ~chunk_size pieces means the pool always has enough work items to keep
    every core busy, regardless of how the 15 GB is distributed across
    the 9 files."""
    try:
        size = os.path.getsize(filepath)
    except OSError:
        return []
    if size <= chunk_size:
        return [(0, size)]
    offsets = []
    with open(filepath, "rb") as f:
        start = 0
        while start < size:
            end = min(start + chunk_size, size)
            if end < size:
                f.seek(end)
                f.readline()          # snap forward to the next line boundary
                end = f.tell()
            if end <= start:          # safety net against pathological lines
                end = size
            offsets.append((start, end))
            start = end
    return offsets


def _plan_fast_tasks(fast_files, chunk_mb: int):
    """Pre-reads just the header + delimiter + encoding of each file (cheap,
    a few KB) in the main process, then hands out (file, byte-range) work
    items. Fieldnames/delimiter/encoding are computed once here and passed
    to every chunk explicitly — required because each chunk may be executed
    by a different worker PROCESS, which does not share Python-level
    caches with the main process or with each other."""
    chunk_size = max(4, int(chunk_mb or 64)) * 1024 * 1024
    tasks = []
    for f in fast_files:
        try:
            with open(f, "rb") as fh:
                head = fh.readline()
        except Exception:
            continue
        enc = _detect_encoding(f)
        delim = "," if head.count(b",") >= head.count(b";") else ";"
        _DELIM_CACHE[os.path.abspath(f)] = delim
        try:
            fieldnames = next(csv.reader([head.decode(enc, errors="ignore")], delimiter=delim), [])
        except Exception:
            fieldnames = []
        if not fieldnames:
            continue
        for start, end in _split_offsets(f, chunk_size):
            data_start = len(head) if start == 0 else start
            if data_start >= end:
                continue
            tasks.append((f, data_start, end, fieldnames, delim, enc))
    return tasks


def _fast_search_csv(filepath: str, start: int, end: int, fieldnames: list,
                      delim: str, enc: str, target: str, target_norm: str) -> list:
    results = []
    if not target_norm:
        return results

    # FIX: multi-word search was matching the WHOLE query as one literal
    # phrase against the raw csv line ("word1 word2" had to appear
    # contiguous, with exactly one space, in that exact order). CSV lines
    # separate fields with a comma/semicolon, not a space, so the moment
    # the two words the user typed lived in two different columns (or even
    # the same column with extra text between them), that contiguous match
    # could never happen — this is why one word worked but two or three
    # didn't. Now we split the query into tokens and require ALL of them
    # to be present (AND), each one anywhere in the row, in any order.
    tokens_norm = target_norm.split()
    if not tokens_norm:
        return results

    ascii_target = target.isascii()
    body_enc = _body_codec(enc)

    if ascii_target:
        token_bytes_list = [t.encode(body_enc, errors="ignore") for t in tokens_norm]
        variant_res = None
    else:
        token_bytes_list = None
        variant_res = [_variant_byte_pattern(_query_variants(t) or [t], body_enc)
                        for t in tokens_norm]

    try:
        with open(filepath, "rb") as fh:
            fh.seek(start)
            pos = start
            row_num = start        # approximate: chunk-relative, not a
                                    # global line number (counting every
                                    # line from byte 0 of a multi-GB file
                                    # would itself cost as much as the
                                    # search, so it's skipped for speed)
            for raw_line in fh:
                pos += len(raw_line)
                row_num += 1

                if ascii_target:
                    raw_lower = raw_line.lower()
                    hit = all(tb in raw_lower for tb in token_bytes_list)
                else:
                    hit = all(vr and vr.search(raw_line) for vr in variant_res)

                if hit:
                    line = raw_line.decode(body_enc, errors="ignore")
                    line_norm = normalize_text(line)
                    ok = all(t in line_norm for t in tokens_norm)
                    if ok:
                        try:
                            row = next(csv.reader([line], delimiter=delim))
                        except csv.Error:
                            row = None
                        if row is not None:
                            row_cells_norm = [normalize_text(c) for c in row if c]
                            # AND across tokens, but each token may live in
                            # ANY cell of the row (not necessarily the same one)
                            cell_ok = all(any(t in c for c in row_cells_norm) for t in tokens_norm)
                            if cell_ok:
                                row_dict = dict(zip(fieldnames, row))
                                row_dict["__source__"] = os.path.basename(filepath)
                                row_dict["__row__"] = f"~{row_num}"
                                results.append(row_dict)

                if pos >= end:
                    break
    except Exception:
        pass
    return results


def _fast_search_wrap(args):
    filepath, start, end, fieldnames, delim, enc, target, target_norm = args
    rows = _fast_search_csv(filepath, start, end, fieldnames, delim, enc, target, target_norm)
    return pd.DataFrame(rows) if rows else None

# ══════════════════════════════════════════════════════════════
#  FULL ENGINE  (pandas)
# ══════════════════════════════════════════════════════════════
def read_file_safe(f):
    ext = os.path.splitext(f)[1].lower()
    try:
        if ext in (".db", ".sqlite"):
            conn   = sqlite3.connect(f"file:{f}?mode=ro", uri=True)
            tables = [r[0] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table';")]
            parts  = []
            for t in tables:
                df_t = pd.read_sql_query(f"SELECT * FROM [{t}]", conn, dtype=str)
                df_t["__table__"] = t; parts.append(df_t)
            conn.close()
            if parts:
                df = pd.concat(parts, ignore_index=True)
                df["__source__"] = os.path.basename(f); return df
        elif ext == ".xlsx":
            xl = pd.ExcelFile(f, engine="openpyxl"); parts = []
            for sh in xl.sheet_names:
                df_s = pd.read_excel(xl, sheet_name=sh, dtype=str)
                df_s["__sheet__"] = sh; parts.append(df_s)
            if parts:
                df = pd.concat(parts, ignore_index=True)
                df["__source__"] = os.path.basename(f); return df
        elif ext == ".json":
            try:    df = pd.read_json(f, dtype=str)
            except Exception: df = pd.read_json(f, lines=True, dtype=str)
            df["__source__"] = os.path.basename(f); return df
    except Exception:
        pass
    return None

# ══════════════════════════════════════════════════════════════
#  PARALLEL RUNNER
# ══════════════════════════════════════════════════════════════
class _Cancelled(Exception):
    pass

def _start_stop_listener(stop_event):
    if not sys.stdin or not sys.stdin.isatty(): return None

    def _unix_listen():
        import termios, tty, select
        fd  = sys.stdin.fileno()
        old = termios.tcgetattr(fd)
        try:
            tty.setcbreak(fd)
            while not stop_event.is_set():
                r, _, _ = select.select([sys.stdin], [], [], 0.2)
                if r:
                    ch = sys.stdin.read(1)
                    if ch and (ch.lower() == "q" or ch == "\x1b"):
                        stop_event.set(); break
        except Exception:
            pass
        finally:
            try: termios.tcsetattr(fd, termios.TCSADRAIN, old)
            except Exception: pass

    def _win_listen():
        import msvcrt
        while not stop_event.is_set():
            if msvcrt.kbhit():
                ch = msvcrt.getch()
                try:    ch = ch.decode(errors="ignore")
                except: ch = ""
                if ch and (ch.lower() == "q" or ch == "\x1b"):
                    stop_event.set(); break
            time.sleep(0.2)

    t = threading.Thread(target=(_win_listen if os.name == "nt" else _unix_listen), daemon=True)
    t.start(); return t


def _run_parallel(fn_args: list, label: str, colour: str, out_list: list = None,
                   use_process: bool = False) -> list:
    """Runs fn_args=[(fn, arg), ...] with a live progress bar.
    Results are appended directly into out_list (created if not given) so that
    partial results survive even if the run is cancelled mid-way — previously
    a cancellation raised before the list could be returned to the caller,
    silently discarding everything that had already been found.

    FIX (speed): use_process=True runs the pool as separate OS processes
    (ProcessPoolExecutor) instead of threads. Threads share one GIL, so the
    CPU-bound part of line scanning (string compares, csv parsing) was
    effectively running on a single core no matter how many "threads" were
    configured — this is why 9 files could take minutes even with 12 cores
    reported. Processes have no shared GIL, so scanning N files genuinely
    uses N cores in parallel."""
    w        = _workers()
    results  = out_list if out_list is not None else []
    stop     = threading.Event()
    listener = _start_stop_listener(stop)
    Executor = ProcessPoolExecutor if use_process else ThreadPoolExecutor

    console.print(f"  [{C_DIM}]Q / Ctrl+C to cancel[/{C_DIM}]")
    try:
        try:
            with Progress(
                SpinnerColumn(spinner_name="dots2", style=f"bold {colour}"),
                TextColumn(f"[bold {colour}]{{task.description}}"),
                BarColumn(bar_width=None, complete_style=colour,
                          finished_style=C_GREEN, pulse_style=f"dim {colour}"),
                TaskProgressColumn(style=f"bold {C_YELLOW}"),
                MofNCompleteColumn(),
                TimeElapsedColumn(),
                console=console, expand=True, refresh_per_second=20,
            ) as pg:
                task = pg.add_task(label, total=len(fn_args))
                with Executor(max_workers=w) as ex:
                    futs    = {ex.submit(fn, arg): arg for fn, arg in fn_args}
                    pending = set(futs.keys())
                    while pending:
                        if stop.is_set():
                            for f in pending: f.cancel()
                            raise _Cancelled()
                        done, pending = wait(pending, timeout=0.05,
                                             return_when=FIRST_COMPLETED)
                        for fut in done:
                            try:
                                r = fut.result()
                                if r is not None: results.append(r)
                            except Exception:
                                pass
                            pg.advance(task)
        except KeyboardInterrupt:
            stop.set(); raise _Cancelled()
    except _Cancelled:
        stop.set()
        console.print(f"\n  [bold {C_YELLOW}]⚡ Cancelled — partial results kept.[/bold {C_YELLOW}]")
        raise
    finally:
        stop.set()
        if listener: listener.join(timeout=0.5)

    return results

# ══════════════════════════════════════════════════════════════
#  CACHE  — FIX: _write_pkl rebuilds blob THEN pickles,
#           manage_cache always writes meta after success
# ══════════════════════════════════════════════════════════════
_META_COLS = {"__source__", "__sheet__", "__table__"}
_BLOB_COL  = "__blob__"

def _build_blob(df):
    cols = [c for c in df.columns if c not in _META_COLS and c != _BLOB_COL]
    if not cols:
        df[_BLOB_COL] = ""; return df
    combo = df[cols[0]].fillna("").astype(str)
    for c in cols[1:]:
        combo = combo + " " + df[c].fillna("").astype(str)
    df[_BLOB_COL] = combo.map(normalize_text)
    return df

def _meta_path():
    return os.path.join(config["cache_dir"], "meta.json")

def _fsig(f):
    try:
        s = os.stat(f); return f"{s.st_size}:{s.st_mtime_ns}"
    except Exception:
        return "0:0"

def _meta_sig(files):
    raw = "|".join(f"{f}:{_fsig(f)}" for f in files).encode()
    return hashlib.md5(raw).hexdigest()

def cache_valid(files):
    mp = _meta_path()
    if not os.path.isdir(config["cache_dir"]) or not os.path.exists(mp):
        return False
    try:
        m = json.loads(open(mp, encoding="utf-8").read())
        return (m.get("sig")  == _meta_sig(files)
            and m.get("path") == os.path.abspath(config["target_dir"])
            and m.get("exts") == sorted(_active_exts()))
    except Exception:
        return False

def _pkl(f):
    return os.path.join(config["cache_dir"],
                        hashlib.md5(os.path.abspath(f).encode()).hexdigest() + ".pkl")

# FIX: receives a single filepath string (matches _run_parallel signature)
def _write_pkl(filepath: str) -> bool:
    try:
        df = read_file_safe(filepath)
        if df is None:
            return True          # nothing to cache — still counts as done
        df = _build_blob(df)
        pkl_path = _pkl(filepath)
        df.to_pickle(pkl_path)
    except Exception:
        pass                     # silently skip unreadable files
    return True                  # always truthy so progress advances

def manage_cache(full_files) -> bool:
    if not full_files or not config["use_cache"]:
        return True
    if cache_valid(full_files):
        console.print(
            f"  [bold {C_LIME}]✔ Cache up-to-date[/bold {C_LIME}]"
            f"  [{C_DIM}]({len(full_files)} file(s))[/{C_DIM}]"
        )
        return True

    os.makedirs(config["cache_dir"], exist_ok=True)
    console.print(
        f"\n  [bold {C_ORANGE}]⚡ Building cache[/bold {C_ORANGE}]"
        f"  [{C_DIM}]{len(full_files)} file(s) · {_workers()} threads[/{C_DIM}]"
    )
    try:
        # FIX: _write_pkl takes a single string arg
        _run_parallel(
            [(_write_pkl, f) for f in full_files],
            label="Caching files",
            colour=C_ORANGE,
        )
    except _Cancelled:
        console.print(f"  [bold {C_YELLOW}]Cache build cancelled.[/bold {C_YELLOW}]\n")
        return False

    # FIX: always write meta after a successful (even partial) build
    try:
        with open(_meta_path(), "w", encoding="utf-8") as mf:
            json.dump({
                "sig":  _meta_sig(full_files),
                "path": os.path.abspath(config["target_dir"]),
                "exts": sorted(_active_exts()),
                "ts":   time.time(),
            }, mf)
    except Exception:
        pass

    n_f, csize = _cache_stats()
    console.print(
        f"  [bold {C_LIME}]✔ Cache ready[/bold {C_LIME}]"
        f"  [{C_DIM}]{n_f} file(s) · {csize}[/{C_DIM}]\n"
    )
    return True

# ══════════════════════════════════════════════════════════════
#  BLOB SEARCH  (pandas engine)
# ══════════════════════════════════════════════════════════════
def _search_df(df, term_norm):
    if df is None or df.empty: return None
    if _BLOB_COL not in df.columns: df = _build_blob(df)
    # FIX: multi-word queries were matched as ONE literal phrase against the
    # blob ("word1 word2" had to appear contiguous, with exactly one space,
    # in that exact order). Since the blob glues columns together with a
    # single space, this broke the moment another column sat between the
    # two words the user was actually looking for. Split into tokens and
    # require every token to appear SOMEWHERE in the blob (AND, any order,
    # any position) — this is what "search for two/three words" should mean.
    tokens = term_norm.split()
    if not tokens:
        return None
    hit = pd.Series(True, index=df.index)
    for t in tokens:
        hit &= df[_BLOB_COL].str.contains(t, regex=False, na=False)
    found = df.loc[hit, [c for c in df.columns if c != _BLOB_COL]]
    return found if not found.empty else None

_MEM_CACHE: dict = {}
_MEM_CACHE_LOCK = threading.Lock()

def _load_search_full(f, term_norm):
    sig = _fsig(f)
    with _MEM_CACHE_LOCK:
        cached = _MEM_CACHE.get(f)
    if cached and cached[0] == sig:
        df = cached[1]
    else:
        if config["use_cache"]:
            try:    df = pd.read_pickle(_pkl(f))
            except Exception: df = read_file_safe(f)
        else:
            df = read_file_safe(f)
        if df is not None:
            if _BLOB_COL not in df.columns: df = _build_blob(df)
            with _MEM_CACHE_LOCK:
                _MEM_CACHE[f] = (sig, df)
        else:
            return None
    return _search_df(df, term_norm)

def _load_search_full_wrap(args):
    f, term_norm = args
    return _load_search_full(f, term_norm)

# ══════════════════════════════════════════════════════════════
#  RESULTS TABLE
# ══════════════════════════════════════════════════════════════
def _cell(val) -> str:
    s = "" if (val is None or (isinstance(val, float) and pd.isna(val))) else str(val)
    return fix_rtl(s) if _HAS_BIDI else s

def _print_results(final: pd.DataFrame):
    special = [c for c in ("__source__", "__sheet__", "__table__", "__row__")
               if c in final.columns]
    other   = [c for c in final.columns if c not in set(special)]
    display = (special + other)[:10]
    _META_STYLE = {
        "__source__": f"bold {C_PINK}",
        "__sheet__":  f"bold {C_ORANGE}",
        "__table__":  f"bold {C_PURPLE}",
        "__row__":    C_DIM,
    }
    tbl = Table(show_header=True, header_style=f"bold {C_CYAN_LT}",
                border_style=C_CYAN_DK, box=box.SIMPLE_HEAD,
                row_styles=[C_WHITE, f"dim {C_WHITE}"],
                show_edge=True, padding=(0, 1), expand=False,
                title=f"[bold {C_CYAN_LT}]◈ Results[/bold {C_CYAN_LT}]",
                title_style="", title_justify="left",
                caption=f"[{C_DIM}]row 1–{min(len(final),50)} of {len(final):,}[/{C_DIM}]",
                caption_justify="left")
    for col in display:
        # FIX: was overflow="fold", which chops long cells mid-word. For
        # bidi-reshaped Persian/Arabic text that corrupts the glyph joins
        # and makes words look scrambled. "ellipsis" + no_wrap truncates
        # cleanly instead of breaking a shaped string across lines.
        tbl.add_column(col, style=_META_STYLE.get(col, ""), max_width=40,
                       overflow="ellipsis", no_wrap=True)
    LIMIT = 50
    for _, row in final.head(LIMIT).iterrows():
        tbl.add_row(*[_cell(row[c]) for c in display])
    console.print(tbl)
    if len(final) > LIMIT:
        console.print(f"  [{C_DIM}]… {len(final)-LIMIT:,} more rows (export to CSV for full view).[/{C_DIM}]\n")

# ══════════════════════════════════════════════════════════════
#  RUN SEARCH
# ══════════════════════════════════════════════════════════════
def run_search():
    print_header()
    files = get_files(config["target_dir"])
    if not files:
        _err("No files found in target directory.")
        questionary.press_any_key_to_continue(style=_STYLE).ask()
        return

    fast_files = [f for f in files if os.path.splitext(f)[1].lower() in _FAST_EXTS]
    full_files = [f for f in files if os.path.splitext(f)[1].lower() not in _FAST_EXTS]

    if full_files:
        ok = manage_cache(full_files)
        if not ok:
            go = questionary.confirm(
                "  Cache incomplete — continue with live reads?",
                default=True, style=_STYLE).ask()
            if not go: return

    total_size = sum(os.path.getsize(f) for f in files if os.path.exists(f))
    console.print(Panel(
        f"  [bold {C_LIME}]⚡ Fast  [/bold {C_LIME}][{C_WHITE}]{len(fast_files)}[/{C_WHITE}]  [{C_DIM}](csv/txt)[/{C_DIM}]\n"
        f"  [bold {C_PURPLE}]⊞  Full  [/bold {C_PURPLE}][{C_WHITE}]{len(full_files)}[/{C_WHITE}]  [{C_DIM}](xlsx/json/db)[/{C_DIM}]",
        title=f"[bold {C_CYAN_LT}]◆ FILES — {len(files)} total · {total_size/1_073_741_824:.2f} GB[/bold {C_CYAN_LT}]",
        border_style=C_CYAN_DK, box=box.ROUNDED, padding=(0, 1)
    ))
    console.print()

    term = questionary.text("  🔎  Search term (empty → cancel):", style=_STYLE, qmark="").ask()
    if not term or not term.strip():
        return
    term_stripped = term.strip()
    term_norm = normalize_text(term_stripped)
    if not term_norm:
        return

    t_start = time.perf_counter()
    raw = []

    if fast_files:
        chunk_mb = config.get("chunk_mb", 64)
        tasks = _plan_fast_tasks(fast_files, chunk_mb)
        total_bytes = sum(os.path.getsize(f) for f in fast_files if os.path.exists(f))
        console.print(
            f"\n  [bold {C_LIME}]⚡ Fast Engine[/bold {C_LIME}]"
            f"  [{C_DIM}]{len(fast_files)} file(s) · {len(tasks)} chunk(s)"
            f" · {total_bytes/1_073_741_824:.2f} GB[/{C_DIM}]"
        )
        try:
            _run_parallel(
                [(_fast_search_wrap, (f, s, e, fn, dl, en, term_stripped, term_norm))
                 for (f, s, e, fn, dl, en) in tasks],
                label=f'Scanning "{term_stripped}"', colour=C_GREEN, out_list=raw,
                use_process=True)
        except _Cancelled:
            console.print(f"  [bold {C_YELLOW}]Fast scan cancelled.[/bold {C_YELLOW}]\n")

    if full_files:
        console.print(f"\n  [bold {C_PURPLE}]⊞  Full Engine[/bold {C_PURPLE}]  [{C_DIM}]{len(full_files)} file(s)[/{C_DIM}]")
        try:
            _run_parallel(
                [(_load_search_full_wrap, (f, term_norm)) for f in full_files],
                label=f'Searching "{term_stripped}"', colour=C_PURPLE, out_list=raw)
        except _Cancelled:
            console.print(f"  [bold {C_YELLOW}]Full search cancelled.[/bold {C_YELLOW}]\n")

    elapsed = time.perf_counter() - t_start

    if not raw:
        _err(f'No matches for  [bold {C_YELLOW}]"{term_stripped}"[/bold {C_YELLOW}]  '
             f'[{C_DIM}]({elapsed:.2f}s)[/{C_DIM}]')
        questionary.press_any_key_to_continue(style=_STYLE).ask()
        return

    final = pd.concat(raw, ignore_index=True)
    scanned_bytes = sum(os.path.getsize(f) for f in files if os.path.exists(f))
    throughput = (scanned_bytes / 1_048_576 / elapsed) if elapsed > 0 else 0
    console.print(Panel(
        f"  [bold {C_LIME}]✔  {len(final):,} match(es)[/bold {C_LIME}]"
        f"  [{C_DIM}]{len(raw)} chunk(s) with hits · {elapsed:.2f}s[/{C_DIM}]\n"
        f"  {_speed_badge(throughput)}  [{C_DIM}]{throughput:,.0f} MB/s[/{C_DIM}]",
        border_style=C_LIME, box=box.ROUNDED, padding=(0, 1)
    ))
    console.print()
    _print_results(final)

    if questionary.confirm("  Export results to CSV?", default=False, style=_STYLE).ask():
        fname = questionary.text("  Filename:", default="results.csv", style=_STYLE).ask()
        if fname and fname.strip():
            out_path = fname.strip()
            final.to_csv(out_path, index=False, encoding="utf-8-sig")
            _ok(f"Saved →  [{C_WHITE}]{os.path.abspath(out_path)}[/{C_WHITE}]")

    questionary.press_any_key_to_continue(style=_STYLE).ask()

# ══════════════════════════════════════════════════════════════
#  EXT SETTINGS
#  FIX: custom checkbox now correctly saves the selected list;
#       preset menu shows active preset with ✔ marker
# ══════════════════════════════════════════════════════════════
_EXT_PRESETS = {
    "csv_only":  [".csv"],
    "csv_txt":   [".csv", ".txt"],
    "csv_excel": [".csv", ".xlsx"],
    "csv_db":    [".csv", ".db", ".sqlite"],
    "all":       sorted(ALL_EXTS),
}

def _ext_settings():
    while True:
        print_header()
        console.print(Panel(
            f"  [{C_DIM}]Default: CSV only  (⚡ fastest — line-by-line engine)[/{C_DIM}]\n"
            f"  [{C_DIM}]xlsx/json/db use the pandas engine (slower, more thorough).[/{C_DIM}]",
            title=f"[bold {C_CYAN_LT}]FILE EXTENSION FILTER[/bold {C_CYAN_LT}]",
            border_style=C_CYAN_DK, box=box.ROUNDED, padding=(0, 1)
        ))
        console.print()

        current = frozenset(_active_exts())

        def _mark(key):
            return "✔" if frozenset(_EXT_PRESETS[key]) == current else " "

        # FIX: mark custom if current doesn't match any preset
        custom_mark = " " if any(frozenset(v) == current for v in _EXT_PRESETS.values()) else "✔"

        choice = questionary.select(
            "  Select preset:",
            choices=[
                questionary.Choice(f"  {_mark('csv_only')}  ⚡  CSV only            [Fastest]",                    value="csv_only"),
                questionary.Choice(f"  {_mark('csv_txt')}   ⚡  CSV + TXT           [Fast]",                        value="csv_txt"),
                questionary.Choice(f"  {_mark('csv_excel')}    CSV + Excel         [.csv .xlsx]",                   value="csv_excel"),
                questionary.Choice(f"  {_mark('csv_db')}    CSV + DB             [.csv .db .sqlite]",               value="csv_db"),
                questionary.Choice(f"  {_mark('all')}    All Formats         [csv txt xlsx json db sqlite]",        value="all"),
                questionary.Choice(f"  {custom_mark}  Custom — choose manually …",                                  value="custom"),
                questionary.Choice("      ↩  Back",                                                                  value="back"),
            ],
            style=_STYLE, qmark="", instruction="[↑/↓ · Enter]",
        ).ask()

        if choice is None or choice == "back":
            return

        if choice in _EXT_PRESETS:
            config["active_exts"] = list(_EXT_PRESETS[choice])
            save_config(config)
            _ok(f"Saved:  [{C_WHITE}]{', '.join(_EXT_PRESETS[choice])}[/{C_WHITE}]", pause=0.8)
            return

        # ── Custom checkbox ───────────────────────────────────────
        # FIX: re-read current exts at this exact moment, not stale
        current_now = _active_exts()
        choices_custom = [
            questionary.Choice(
                title=f"  {e.lstrip('.').upper():<8}  {'⚡ fast' if e in _FAST_EXTS else '⊞  pandas'}",
                value=e,
                checked=(e in current_now),   # FIX: uses live set, not cached frozenset
            )
            for e in sorted(ALL_EXTS)
        ]
        selected = questionary.checkbox(
            "  Toggle extensions  (Space = toggle, Enter = confirm):",
            choices=choices_custom,
            style=_STYLE, qmark="", instruction="[Space · ↑/↓ · Enter]",
        ).ask()

        if selected is None:
            # user pressed Ctrl+C — go back to preset menu
            continue

        if not selected:
            _warn("At least one extension required — no change.", pause=1.0)
            continue

        # FIX: store the actual selected list (was sometimes storing old value)
        new_exts = sorted(set(selected))
        config["active_exts"] = new_exts
        save_config(config)
        _ok(f"Saved:  [{C_WHITE}]{', '.join(new_exts)}[/{C_WHITE}]", pause=0.8)
        return

# ══════════════════════════════════════════════════════════════
#  SETTINGS
#  FIX: no duplicate CONFIGURATION panel — only SYSTEM STATUS
#       appears in print_header(); settings shows only the menu
# ══════════════════════════════════════════════════════════════
def open_settings():
    while True:
        print_header()   # contains SYSTEM STATUS — no extra panel needed
        # Show a compact current-values strip instead of a full duplicate panel
        exts_lbl = ", ".join(sorted(_active_exts()))
        tog       = "ON → OFF" if config["use_cache"] else "OFF → ON"
        w_lbl     = "UNLIMITED" if not config.get("max_workers") else str(config["max_workers"])

        console.print(Rule(f"[bold {C_CYAN_DK}]  Settings  [/bold {C_CYAN_DK}]", style=C_CYAN_DK))
        console.print()

        choice = questionary.select(
            "Choose option:",
            choices=[
                questionary.Choice(f"  ⊞  File extensions              [{exts_lbl}]",  value="exts"),
                questionary.Choice(f"  ⊙  Toggle cache                 [{tog}]",         value="cache"),
                questionary.Choice(f"  ⚙  Thread limit                 [{w_lbl}]",       value="threads"),
                questionary.Choice(f"  🧩  Chunk size (fast engine)     [{config.get('chunk_mb',64)} MB]", value="chunk"),
                questionary.Choice("  📁  Change target directory",                        value="target"),
                questionary.Choice("  💾  Change cache directory",                         value="cachedir"),
                questionary.Choice("  ↩  Back",                                            value="back"),
            ],
            style=_STYLE, qmark="⚙", instruction="[↑/↓ · Enter]",
        ).ask()

        if choice is None or choice == "back":
            break

        elif choice == "exts":
            _ext_settings()

        elif choice == "cache":
            config["use_cache"] = not config["use_cache"]
            save_config(config)
            state = "ENABLED" if config["use_cache"] else "DISABLED"
            _ok(f"Cache {state}", pause=0.7)

        elif choice == "threads":
            presets = (["0 — UNLIMITED (all cores)"]
                       + [str(n) for n in [1,2,4,8,16,32] if n <= _CPU*2]
                       + ["Custom …", "↩ Cancel"])
            val = questionary.select(
                f"  Thread limit  (machine: {_CPU} cores):",
                choices=presets, style=_STYLE, qmark="⚙",
            ).ask()
            if val is None or val == "↩ Cancel": continue
            elif val.startswith("0"):  config["max_workers"] = None
            elif val == "Custom …":
                raw = questionary.text("  Enter number (0 = unlimited):", style=_STYLE).ask()
                if raw and raw.strip().isdigit():
                    n = int(raw.strip())
                    config["max_workers"] = None if n == 0 else max(1, n)
            else:
                config["max_workers"] = int(val.split()[0])
            save_config(config)

        elif choice == "chunk":
            presets = ["16", "32", "64", "128", "256", "Custom …", "↩ Cancel"]
            val = questionary.select(
                "  Chunk size in MB — smaller = more parallel chunks, more overhead;"
                " larger = fewer chunks, less overhead:",
                choices=presets, style=_STYLE, qmark="🧩",
            ).ask()
            if val is None or val == "↩ Cancel": continue
            elif val == "Custom …":
                raw_v = questionary.text("  Enter MB (e.g. 96):", style=_STYLE).ask()
                if raw_v and raw_v.strip().isdigit():
                    config["chunk_mb"] = max(4, int(raw_v.strip()))
            else:
                config["chunk_mb"] = int(val)
            save_config(config)
            _ok("Chunk size updated.", pause=0.5)

        elif choice == "target":
            new = questionary.path("  Target directory:", default=config["target_dir"], style=_STYLE).ask()
            if new and new.strip():
                config["target_dir"] = new.strip(); save_config(config)
                _ok("Target updated.", pause=0.5)

        elif choice == "cachedir":
            new = questionary.path("  Cache directory:", default=config["cache_dir"], style=_STYLE).ask()
            if new and new.strip():
                os.makedirs(new.strip(), exist_ok=True)
                config["cache_dir"] = new.strip(); save_config(config)
                _ok("Cache dir updated.", pause=0.5)

# ══════════════════════════════════════════════════════════════
#  MAIN MENU
# ══════════════════════════════════════════════════════════════
def main():
    if not _HAS_BIDI:
        # FIX: without arabic_reshaper + python-bidi, Persian/Arabic text is
        # printed in raw logical order with unjoined letters — this is the
        # "جدا و بی‌ترتیب" (scrambled) look. Tell the user how to fix it
        # instead of silently showing broken text.
        _warn("Persian/Arabic text will look scrambled.")
        _dim("Run:  pip install arabic-reshaper python-bidi")
        time.sleep(1.5)
    while True:
        print_header()
        console.print(Rule(style=C_CYAN_DK))
        choice = questionary.select(
            "  What would you like to do?",
            choices=[
                questionary.Choice("  🔍  Search databases",  value="search"),
                questionary.Choice("  ⚙️   Settings & paths",  value="settings"),
                questionary.Choice("  🗑️   Clear cache",       value="clear"),
                questionary.Choice("  ❌   Exit",              value="exit"),
            ],
            style=_STYLE, instruction="[↑/↓ · Enter]", qmark="❯",
        ).ask()

        if choice is None or choice == "exit":
            os.system("cls" if os.name == "nt" else "clear")   # FIX: OS-conditional clear
            console.print()
            console.print(Align.center(_gradient_rule(width=48, ornament="✦")))
            console.print(Align.center(
                f"[bold {C_PINK}]  See you soon, @GreatSaadi 🤠[/bold {C_PINK}]",
                vertical="middle"))
            console.print(Align.center(_gradient_rule(width=48, ornament="✦")))
            console.print()
            break

        elif choice == "search":
            run_search()

        elif choice == "settings":
            open_settings()

        elif choice == "clear":
            n_f, csize = _cache_stats()
            if n_f == 0:
                _dim("Cache is already empty.")
                time.sleep(0.8); continue
            if questionary.confirm(
                f"  Delete {n_f} cached file(s) ({csize})?",
                default=False, style=_STYLE).ask():
                if os.path.isdir(config["cache_dir"]):
                    shutil.rmtree(config["cache_dir"])
                _MEM_CACHE.clear()
                _ok("Cache cleared.", pause=0.8)

if __name__ == "__main__":
    # CRITICAL for frozen (.exe) builds on Windows: ProcessPoolExecutor
    # spawns workers by re-launching this same executable. Without
    # freeze_support() called FIRST, the frozen bootloader has no way to
    # tell "I am a worker process" from "I am a fresh app launch" — so
    # every worker re-runs the ENTIRE app (menu, prompts, everything)
    # instead of quietly waiting for work. That's what caused the
    # duplicated SYSTEM STATUS panels and the massive slowdown: each of
    # the 12 worker processes was trying to render/run its own copy of
    # the interactive app on top of the real one. This must be the very
    # first line executed — before any UI, before main().
    import multiprocessing
    multiprocessing.freeze_support()

    if os.name == "nt":
        os.system("color")
    main()