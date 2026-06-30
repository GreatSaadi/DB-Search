"""
╔══════════════════════════════════════════════════════════════╗
║        DATABASE SEARCHER  v1.0.0  —  @GreatSaadi             ║
║              https://github.com/GreatSaadi                   ║
║                       DB Searcher                            ║
╚══════════════════════════════════════════════════════════════╝
"""

import os, sys, sqlite3, hashlib, json, time, shutil, threading, unicodedata, re, csv
import pandas as pd
import platform
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
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
    def fix_rtl(text: str) -> str:
        if not text or not isinstance(text, str): return text
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
#  CONFIG LOAD / SAVE
# ══════════════════════════════════════════════════════════════
def load_config():
    d = {
        "target_dir":  "./",
        "use_cache":   False,
        "cache_dir":   os.path.join(SCRIPT_DIR, "db_searcher_cache"),
        "max_workers": None,
        "active_exts": [".csv"],
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
    # author strip
    console.print(
        f"  [bold {C_CYAN}]v1.0[/bold {C_CYAN}]"
        f"  [{C_DIM}]·[/{C_DIM}]"
        f"  [italic bold {C_CYAN_LT}]@GreatSaadi[/italic bold {C_CYAN_LT}]"
        f"  [{C_DIM}]·[/{C_DIM}]"
        f"  [{C_CYAN_LT}]github.com/GreatSaadi[/{C_CYAN_LT}]"
    )
    console.print(f"  [{C_WHITE}]" + "─" * 77 + f"[/{C_WHITE}]")
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

    rows = "\n".join([
        f"  [bold {C_CYAN_DK}]TARGET  [/bold {C_CYAN_DK}] [{C_WHITE}]{tgt}[/{C_WHITE}]",
        f"  [bold {C_CYAN_DK}]CACHE   [/bold {C_CYAN_DK}] {cache_val}",
        f"  [bold {C_CYAN_DK}]CDIR    [/bold {C_CYAN_DK}] [{C_DIM}]{cdir}[/{C_DIM}]",
        f"  [bold {C_CYAN_DK}]THREADS [/bold {C_CYAN_DK}] {tbar}  {w_label}  [{C_DIM}]({_CPU} cores)[/{C_DIM}]",
        f"  [bold {C_CYAN_DK}]EXTS    [/bold {C_CYAN_DK}] {_ext_badge()}",
    ])
    return Panel(rows,
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
def _fast_search_csv(filepath: str, target: str) -> list:
    results = []
    for enc in ENCODINGS:
        try:
            with open(filepath, "r", encoding=enc, errors="ignore") as fh:
                sample = fh.read(4096); fh.seek(0)
                delim  = "," if sample.count(",") >= sample.count(";") else ";"
                reader = csv.DictReader(fh, delimiter=delim)
                fnames = reader.fieldnames or []
                for row_num, row in enumerate(reader, 2):
                    for field in fnames:
                        if field and field in row and row[field]:
                            if target in str(row[field]):
                                r = dict(row)
                                r["__source__"] = os.path.basename(filepath)
                                r["__row__"]    = str(row_num)
                                results.append(r); break
            break
        except Exception:
            continue
    return results

def _fast_search_wrap(args):
    filepath, target = args
    rows = _fast_search_csv(filepath, target)
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


def _run_parallel(fn_args: list, label: str, colour: str, out_list: list = None) -> list:
    """Runs fn_args=[(fn, arg), ...] in a thread pool with a live progress bar.
    Results are appended directly into out_list (created if not given) so that
    partial results survive even if the run is cancelled mid-way — previously
    a cancellation raised before the list could be returned to the caller,
    silently discarding everything that had already been found."""
    w        = _workers()
    results  = out_list if out_list is not None else []
    stop     = threading.Event()
    listener = _start_stop_listener(stop)

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
                with ThreadPoolExecutor(max_workers=w) as ex:
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
    hit   = df[_BLOB_COL].str.contains(term_norm, regex=False, na=False)
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
                show_edge=True, padding=(0, 1), expand=False)
    for col in display:
        tbl.add_column(col, style=_META_STYLE.get(col, ""), max_width=35,
                       overflow="fold", no_wrap=(col in _META_STYLE))
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
        console.print(f"\n  [bold {C_RED}]✕  No files found in target directory.[/bold {C_RED}]")
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

    console.print(Panel(
        f"  [bold {C_LIME}]⚡ Fast  [/bold {C_LIME}][{C_WHITE}]{len(fast_files)}[/{C_WHITE}]  [{C_DIM}](csv/txt)[/{C_DIM}]\n"
        f"  [bold {C_PURPLE}]⊞  Full  [/bold {C_PURPLE}][{C_WHITE}]{len(full_files)}[/{C_WHITE}]  [{C_DIM}](xlsx/json/db)[/{C_DIM}]",
        title=f"[bold {C_CYAN_LT}]FILES — {len(files)} total[/bold {C_CYAN_LT}]",
        border_style=C_CYAN_DK, box=box.ROUNDED, padding=(0, 1)
    ))
    console.print()

    term = questionary.text("  Search term (empty → cancel):", style=_STYLE).ask()
    if term is None or not term.strip(): return
    term_stripped = term.strip()
    term_norm     = normalize_text(term_stripped)
    if not term_norm: return

    t_start = time.perf_counter()
    raw = []

    if fast_files:
        console.print(f"\n  [bold {C_LIME}]⚡ Fast Engine[/bold {C_LIME}]  [{C_DIM}]{len(fast_files)} file(s)[/{C_DIM}]")
        try:
            _run_parallel(
                [(_fast_search_wrap, (f, term_stripped)) for f in fast_files],
                label=f'Scanning "{term_stripped}"', colour=C_GREEN, out_list=raw)
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
        console.print(
            f"\n  [bold {C_RED}]✕  No matches for[/bold {C_RED}]"
            f"  [{C_YELLOW}]\"{term_stripped}\"[/{C_YELLOW}]  [{C_DIM}]{elapsed:.2f}s[/{C_DIM}]")
        questionary.press_any_key_to_continue(style=_STYLE).ask()
        return

    final = pd.concat(raw, ignore_index=True)
    console.print(Panel(
        f"  [bold {C_LIME}]✔  {len(final):,} match(es)[/bold {C_LIME}]"
        f"  [{C_DIM}]{len(raw)} file(s) · {elapsed:.2f}s[/{C_DIM}]",
        border_style=C_LIME, box=box.ROUNDED, padding=(0, 1)
    ))
    console.print()
    _print_results(final)

    if questionary.confirm("  Export results to CSV?", default=False, style=_STYLE).ask():
        fname = questionary.text("  Filename:", default="results.csv", style=_STYLE).ask()
        if fname and fname.strip():
            out_path = fname.strip()
            final.to_csv(out_path, index=False, encoding="utf-8-sig")
            console.print(f"  [bold {C_LIME}]✔ Saved →[/bold {C_LIME}]  [{C_WHITE}]{os.path.abspath(out_path)}[/{C_WHITE}]")

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
            console.print(
                f"\n  [bold {C_LIME}]✔ Saved:[/bold {C_LIME}]"
                f"  [{C_WHITE}]{', '.join(_EXT_PRESETS[choice])}[/{C_WHITE}]"
            )
            time.sleep(0.8)
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
            console.print(f"\n  [{C_YELLOW}]⚠  At least one extension required — no change.[/{C_YELLOW}]")
            time.sleep(1.0)
            continue

        # FIX: store the actual selected list (was sometimes storing old value)
        new_exts = sorted(set(selected))
        config["active_exts"] = new_exts
        save_config(config)
        console.print(
            f"\n  [bold {C_LIME}]✔ Saved:[/bold {C_LIME}]"
            f"  [{C_WHITE}]{', '.join(new_exts)}[/{C_WHITE}]"
        )
        time.sleep(0.8)
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
            console.print(f"\n  [bold {C_LIME}]✔ Cache {state}[/bold {C_LIME}]")
            time.sleep(0.7)

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

        elif choice == "target":
            new = questionary.path("  Target directory:", default=config["target_dir"], style=_STYLE).ask()
            if new and new.strip():
                config["target_dir"] = new.strip(); save_config(config)
                console.print(f"\n  [bold {C_LIME}]✔ Target updated.[/bold {C_LIME}]"); time.sleep(0.5)

        elif choice == "cachedir":
            new = questionary.path("  Cache directory:", default=config["cache_dir"], style=_STYLE).ask()
            if new and new.strip():
                os.makedirs(new.strip(), exist_ok=True)
                config["cache_dir"] = new.strip(); save_config(config)
                console.print(f"\n  [bold {C_LIME}]✔ Cache dir updated.[/bold {C_LIME}]"); time.sleep(0.5)

# ══════════════════════════════════════════════════════════════
#  MAIN MENU
# ══════════════════════════════════════════════════════════════
def main():
    while True:
        print_header()
        choice = questionary.select(
            "What would you like to do?",
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
            console.print(Align.center(
                f"\n[bold {C_PINK}]  See you soon, @GreatSaadi 🤠[/bold {C_PINK}]\n",
                vertical="middle"))
            break

        elif choice == "search":
            run_search()

        elif choice == "settings":
            open_settings()

        elif choice == "clear":
            n_f, csize = _cache_stats()
            if n_f == 0:
                console.print(f"\n  [{C_DIM}]Cache is already empty.[/{C_DIM}]")
                time.sleep(0.8); continue
            if questionary.confirm(
                f"  Delete {n_f} cached file(s) ({csize})?",
                default=False, style=_STYLE).ask():
                if os.path.isdir(config["cache_dir"]):
                    shutil.rmtree(config["cache_dir"])
                _MEM_CACHE.clear()
                console.print(f"  [bold {C_LIME}]✔ Cache cleared.[/bold {C_LIME}]")
                time.sleep(0.8)

if __name__ == "__main__":
    if os.name == "nt":
        os.system("color")
    main()