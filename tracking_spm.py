import hashlib
import json
import os
import re
import threading
import queue
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Dependencies: pandas, openpyxl, xlrd (untuk file .xls seperti Andon Monthly.xls).

DEFAULT_SPM_FOLDER = r"\\kmj2dataserver\STOCK SUPERMARKET\SUPERMARKET\2026\GEDUNG A\A NOVEMBER 2026"
DEFAULT_PROD_FOLDER = r"\\kmj2dataserver\PPIC\Hasil & Target Prod\HASIL & TARGET PRODUKSI TAHUN 2026\11. HASIL & TARGET PRODUKSI NOVEMBER 2026\HASIL & TARGET 02-30 NOVEMBER 2026"
DEFAULT_ANDON_FILE = r"Andon Monthly.xls"
CACHE_DIR = os.path.join(os.path.expanduser("~"), ".spm_produksi_cache_v7")
MAX_PREVIEW_ROWS = 5000
TREE_BATCH_SIZE = 250
SUPPORTED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm")

OUTPUT_COLUMNS = [
    "TANGGAL",
    "LINE",
    "KT ORDER",
    "STYLE",
    "COLOR",
    "DESCRIPTION",
    "ISI BUNDLE",
    "STATUS",
    "BUNDLING OUT",
    "REQ SPM",
    "LEAD TIME",
    "HASIL PRODUKSI",
]

SPM_HEADER_ALIASES = {
    "LINE": {"LINE"},
    "KT ORDER": {"KJ NO", "KJ # NO", "KT ORDER", "KT NO"},
    "STYLE": {"STYLE"},
    "COLOR": {"COLOUR", "COLOR"},
    "DESCRIPTION": {"DESCRIPTION", "DESC"},
    "ISI BUNDLE": {"ISI BUNDLE", "ISI BDL"},
    "STATUS": {"STATUS", "STAT"},
    "BUNDLING OUT": {"OUT BUND", "OUT BUNDLE", "BUNDLING OUT"},
}

PROD_HEADER_ALIASES = {
    "LINE": {"LINE"},
    "STYLE": {"STYLE"},
}


def norm_header(val) -> str:
    s = str(val or "").strip().upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s).strip()
    return s


def norm_section_name(val) -> str:
    s = str(val or "").strip().lower()
    # Samakan typo "sweing" menjadi "sewing".
    s = re.sub(r"\bsweing\b", "sewing", s)
    # Nama section pada Andon kadang menyisipkan kata "Line".
    # Hilangkan kata tersebut agar "Sweing Line to SuperMarket" cocok
    # dengan target "sewing to supermarket".
    s = re.sub(r"\bline\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_spm_date(filename: str):
    m = re.search(r"KMJ2\s+(\d{8})\s+GEDUNG", filename, re.IGNORECASE)
    if not m:
        m = re.search(r"(\d{8})", filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None


def parse_prod_date(filename: str):
    m = re.search(r"\((\d{2}-\d{2}-\d{2,4})\)", filename)
    if not m:
        return None
    parts = m.group(1).split("-")
    day, month, year = int(parts[0]), int(parts[1]), int(parts[2])
    if year < 100:
        year += 2000
    try:
        return datetime(year, month, day).date()
    except ValueError:
        return None


def parse_input_date(text: str):
    s = (text or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Format tanggal tidak valid: '{s}'. Gunakan DD-MM-YYYY atau YYYY-MM-DD.")


def parse_line_filter(text: str):
    s = (text or "").strip()
    if not s or s.upper() in {"ALL", "*"}:
        return None
    lines = set()
    for part in s.split(","):
        p = part.strip()
        if p:
            lines.add(int(p))
    return lines


def parse_style_filters(text: str) -> list[str]:
    s = (text or "").strip().upper()
    if not s or s in {"ALL", "*"}:
        return []
    return [part.strip() for part in s.split(",") if part.strip()]


def to_number(val):
    if val is None or val == "":
        return 0
    try:
        num = float(val)
        return int(num) if num.is_integer() else round(num, 4)
    except (ValueError, TypeError):
        s = str(val).strip().replace(",", "")
        try:
            num = float(s)
            return int(num) if num.is_integer() else round(num, 4)
        except ValueError:
            return 0


def parse_count_value(val) -> float:
    """Konversi nilai Count dari Andon (termasuk string bertanda kutip/spasi/koma) menjadi float."""
    if val is None or (not isinstance(val, (list, tuple, dict)) and pd.isna(val)):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.upper() in {"-", "—", "N/A", "NA", "NONE", "NULL"}:
        return 0.0
    s = s.replace("'", "").replace('"', "").replace(",", "").replace(" ", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        cleaned = re.sub(r"[^0-9.\-]", "", s)
        try:
            return float(cleaned) if cleaned not in {"", "-", "."} else 0.0
        except ValueError:
            return 0.0


def _normalize_line_value(val):
    if val is None or (not isinstance(val, (list, tuple, dict)) and pd.isna(val)):
        return None
    s = str(val).strip()
    if not s:
        return None
    try:
        num = float(s)
        return int(num) if num.is_integer() else s
    except ValueError:
        return s.upper()


def _normalize_date_value(val):
    if val is None or (not isinstance(val, (list, tuple, dict)) and pd.isna(val)):
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    try:
        parsed = pd.to_datetime(val, errors="coerce")
        if pd.isna(parsed):
            return None
        return parsed.date()
    except Exception:
        return None


def date_in_range(tgl: date | None, start_dt: date | None, end_dt: date | None) -> bool:
    if tgl is None:
        return start_dt is None and end_dt is None
    if start_dt and tgl < start_dt:
        return False
    if end_dt and tgl > end_dt:
        return False
    return True


def extract_style_key(style_text: str) -> str:
    s = str(style_text or "").strip().upper()
    if "/" in s:
        s = s.split("/")[-1].strip()
    return s


def get_cell(row, idx: int | None):
    if idx is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def get_file_cache_paths(folder_path: str, kind: str, fname: str):
    os.makedirs(CACHE_DIR, exist_ok=True)
    folder_hash = hashlib.md5(f"{kind}::{os.path.abspath(folder_path)}".encode("utf-8")).hexdigest()[:12]
    file_hash = hashlib.md5(fname.encode("utf-8")).hexdigest()[:12]
    base = os.path.join(CACHE_DIR, f"{folder_hash}_{file_hash}")
    return f"{base}.pkl", f"{base}.json"


def load_file_cache_if_valid(folder_path: str, kind: str, fname: str, fpath: str) -> pd.DataFrame | None:
    pkl_path, meta_path = get_file_cache_paths(folder_path, kind, fname)
    if not (os.path.exists(pkl_path) and os.path.exists(meta_path)):
        return None
    try:
        stat = os.stat(fpath)
        with open(meta_path, "r", encoding="utf-8") as file:
            meta = json.load(file)
        if meta.get("size") == stat.st_size and meta.get("mtime_ns") == stat.st_mtime_ns:
            return pd.read_pickle(pkl_path)
    except Exception:
        return None
    return None


def save_file_cache(folder_path: str, kind: str, fname: str, fpath: str, df: pd.DataFrame):
    pkl_path, meta_path = get_file_cache_paths(folder_path, kind, fname)
    pkl_tmp = f"{pkl_path}.{threading.get_ident()}.tmp"
    meta_tmp = f"{meta_path}.{threading.get_ident()}.tmp"
    try:
        stat = os.stat(fpath)
        df.to_pickle(pkl_tmp)
        with open(meta_tmp, "w", encoding="utf-8") as file:
            json.dump(
                {
                    "folder": folder_path,
                    "file": fname,
                    "size": stat.st_size,
                    "mtime_ns": stat.st_mtime_ns,
                    "rows": len(df),
                    "cached_at": datetime.now().isoformat(timespec="seconds"),
                },
                file,
            )
        os.replace(pkl_tmp, pkl_path)
        os.replace(meta_tmp, meta_path)
    except Exception:
        for temp_path in (pkl_tmp, meta_tmp):
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except OSError:
                pass


def clear_all_cache() -> int:
    if not os.path.isdir(CACHE_DIR):
        return 0
    removed = 0
    for fn in os.listdir(CACHE_DIR):
        fp = os.path.join(CACHE_DIR, fn)
        try:
            os.remove(fp)
            removed += 1
        except OSError:
            pass
    return removed


def detect_spm_header_and_data_start(ws, max_scan_row: int = 15):
    rows_cache = list(ws.iter_rows(min_row=1, max_row=max_scan_row, values_only=True))
    best_data_start = 10
    best_map = {}

    alias_norm = {
        std: {norm_header(alias) for alias in aliases}
        for std, aliases in SPM_HEADER_ALIASES.items()
    }

    for row_idx, row in enumerate(rows_cache):
        if not row:
            continue
        next_row = rows_cache[row_idx + 1] if row_idx + 1 < len(rows_cache) else None
        max_cols = max(len(row), len(next_row) if next_row else 0)
        found = {}
        needs_second_header_row = False

        for col_idx in range(max_cols):
            h1 = norm_header(row[col_idx] if col_idx < len(row) else "")
            h2 = norm_header(next_row[col_idx] if (next_row and col_idx < len(next_row)) else "")
            combined = f"{h1} {h2}".strip()

            for std_col, aliases in alias_norm.items():
                if std_col in found:
                    continue
                if h1 and h1 in aliases:
                    found[std_col] = col_idx
                elif h2 and h2 in aliases:
                    found[std_col] = col_idx
                    needs_second_header_row = True
                elif combined and combined in aliases:
                    found[std_col] = col_idx
                    needs_second_header_row = True

        data_start = row_idx + (3 if needs_second_header_row else 2)
        if len(found) > len(best_map):
            best_map = found
            best_data_start = data_start

        if {"LINE", "STYLE", "STATUS"}.issubset(found.keys()) and len(found) >= 5:
            return data_start, found

    return best_data_start, best_map


def detect_prod_header_and_packing_cols(ws, max_scan_row: int = 10):
    rows_cache = list(ws.iter_rows(min_row=1, max_row=max_scan_row, values_only=True))
    col_map = {"LINE": 0, "STYLE": 2}
    data_start = 5

    alias_norm = {
        std: {norm_header(a) for a in aliases}
        for std, aliases in PROD_HEADER_ALIASES.items()
    }

    for i, row in enumerate(rows_cache):
        if not row:
            continue
        found = {}
        for c_idx, val in enumerate(row):
            nh = norm_header(val)
            if not nh:
                continue
            for std_col, norms in alias_norm.items():
                if std_col not in found and nh in norms:
                    found[std_col] = c_idx
        if {"LINE", "STYLE"}.issubset(found.keys()):
            col_map.update(found)
            data_start = max(5, i + 4)
            break

    packing_cols = []
    for r in rows_cache[:5]:
        if not r:
            continue
        row_pack = [idx for idx, val in enumerate(r) if norm_header(val) == "PACKING"]
        if row_pack:
            packing_cols = row_pack
            break

    if not packing_cols:
        packing_cols = list(range(7, 25))

    return data_start, col_map, packing_cols


def _parse_spm_excel(fpath: str, tgl: date | None) -> pd.DataFrame:
    file_rows = []
    wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True, keep_links=False)
    try:
        sheet_name = "SPM " if "SPM " in wb.sheetnames else next(
            (s for s in wb.sheetnames if s.strip().upper() == "SPM"), None
        )
        if not sheet_name:
            return pd.DataFrame(columns=[
                "TANGGAL", "LINE", "KT ORDER", "STYLE", "STYLE_KEY",
                "COLOR", "DESCRIPTION", "ISI BUNDLE", "STATUS", "BUNDLING OUT"
            ])
        ws = wb[sheet_name]

        data_start, col_map = detect_spm_header_and_data_start(ws)
        idx_line = col_map.get("LINE", 0)
        idx_kt = col_map.get("KT ORDER", 6)
        idx_style = col_map.get("STYLE", 7)
        idx_color = col_map.get("COLOR", 8)
        idx_desc = col_map.get("DESCRIPTION", 9)
        idx_bundle = col_map.get("ISI BUNDLE", 13)
        idx_status = col_map.get("STATUS", 22)
        idx_out = col_map.get("BUNDLING OUT", 23)
        max_needed_col = max(idx_line, idx_kt, idx_style, idx_color, idx_desc, idx_bundle, idx_status, idx_out) + 1

        current_line = None
        for row in ws.iter_rows(min_row=data_start, max_col=max_needed_col, values_only=True):
            if not row:
                continue

            a_val = get_cell(row, idx_line)
            if a_val is not None and str(a_val).strip() != "":
                try:
                    current_line = int(float(str(a_val).strip()))
                except ValueError:
                    current_line = str(a_val).strip()

            style_raw = str(get_cell(row, idx_style) or "").strip()
            if not style_raw:
                continue

            status_val = str(get_cell(row, idx_status) or "").strip().upper()
            color_val = str(get_cell(row, idx_color) or "").strip()

            file_rows.append({
                "TANGGAL": tgl,
                "LINE": current_line,
                "KT ORDER": str(get_cell(row, idx_kt) or "").strip(),
                "STYLE": style_raw,
                "STYLE_KEY": extract_style_key(style_raw),
                "COLOR": color_val,
                "DESCRIPTION": str(get_cell(row, idx_desc) or "").strip(),
                "ISI BUNDLE": to_number(get_cell(row, idx_bundle)),
                "STATUS": status_val,
                "BUNDLING OUT": to_number(get_cell(row, idx_out)),
            })
    finally:
        wb.close()
    return pd.DataFrame(file_rows)


def _parse_packing_excel(fpath: str, tgl: date | None) -> pd.DataFrame:
    records = []
    wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True, keep_links=False)
    try:
        if "GEDUNG A" not in wb.sheetnames:
            return pd.DataFrame(columns=["TANGGAL", "LINE", "STYLE_KEY", "HASIL PRODUKSI"])
        ws = wb["GEDUNG A"]

        data_start, col_map, packing_cols = detect_prod_header_and_packing_cols(ws)
        idx_line = col_map.get("LINE", 0)
        idx_style = col_map.get("STYLE", 2)
        max_needed_col = max([idx_line, idx_style] + packing_cols) + 1

        current_line = None
        for row in ws.iter_rows(min_row=data_start, max_col=max_needed_col, values_only=True):
            if not row:
                continue

            a_val = get_cell(row, idx_line)
            if a_val is not None and str(a_val).strip() != "":
                try:
                    current_line = int(float(str(a_val).strip()))
                except ValueError:
                    current_line = str(a_val).strip()

            style_raw = str(get_cell(row, idx_style) or "").strip()
            if not style_raw:
                continue

            pack_sum = sum(float(to_number(get_cell(row, c))) for c in packing_cols)
            records.append({
                "TANGGAL": tgl,
                "LINE": current_line,
                "STYLE_KEY": extract_style_key(style_raw),
                "HASIL_PRODUKSI_NUM": pack_sum,
            })
    finally:
        wb.close()

    df_p = pd.DataFrame(records)
    if df_p.empty:
        return pd.DataFrame(columns=["TANGGAL", "LINE", "STYLE_KEY", "HASIL PRODUKSI"])

    grouped = df_p.groupby(["TANGGAL", "LINE", "STYLE_KEY"], as_index=False)["HASIL_PRODUKSI_NUM"].sum()
    grouped["HASIL PRODUKSI"] = grouped["HASIL_PRODUKSI_NUM"].apply(
        lambda v: int(v) if float(v).is_integer() else round(float(v), 2)
    )
    return grouped[["TANGGAL", "LINE", "STYLE_KEY", "HASIL PRODUKSI"]]


def load_req_spm_from_andon(andon_path: str, target_section: str = "sewing to supermarket") -> pd.DataFrame:
    """Ambil Count dan Lead Time Andon untuk baris TROUBLE_NM=Request per tanggal + line."""
    output_cols = ["TANGGAL", "LINE", "REQ SPM", "LEAD TIME"]
    empty_df = pd.DataFrame(columns=output_cols)
    if not andon_path or not os.path.isfile(andon_path):
        return empty_df

    ext = os.path.splitext(andon_path)[1].lower()
    engine = "xlrd" if ext == ".xls" else None
    xl = pd.ExcelFile(andon_path, engine=engine)
    sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
    df_raw = xl.parse(sheet_name)

    required_cols = {"SECTION_NM", "LINE", "TROUBLE_NM"}
    if df_raw.empty or not required_cols.issubset(df_raw.columns):
        return empty_df

    day_count_cols = {}
    day_lead_cols = {}
    source_columns = list(df_raw.columns)
    for col_idx, col_name in enumerate(source_columns):
        match = re.match(r"^DAY\s*(\d{1,2})$", str(col_name).strip(), re.IGNORECASE)
        if not match:
            continue

        day_num = int(match.group(1))
        count_header = str(df_raw.iloc[0][col_name] or "").strip().lower()
        if count_header not in {"", "count"}:
            continue
        day_count_cols[day_num] = col_name

        # Pada export Andon, kolom Lead Time berada tepat setelah kolom Count.
        if col_idx + 1 < len(source_columns):
            lead_col = source_columns[col_idx + 1]
            lead_header = str(df_raw.iloc[0][lead_col] or "").strip().lower()
            if lead_header == "lead time":
                day_lead_cols[day_num] = lead_col

    if not day_count_cols:
        return empty_df

    df_data = df_raw.iloc[1:].copy()
    df_data["SECTION_NM"] = df_data["SECTION_NM"].ffill()
    df_data["LINE"] = df_data["LINE"].ffill()
    df_data["TROUBLE_NM"] = df_data["TROUBLE_NM"].ffill()

    target_sec_norm = norm_section_name(target_section)
    section_mask = df_data["SECTION_NM"].apply(norm_section_name).eq(target_sec_norm)
    request_mask = df_data["TROUBLE_NM"].astype(str).str.strip().str.casefold().eq("request")
    df_req = df_data[section_mask & request_mask].copy()
    if df_req.empty:
        return empty_df

    df_req["LINE_NORM"] = df_req["LINE"].apply(_normalize_line_value)
    df_req = df_req[df_req["LINE_NORM"].notna()]

    records = []
    for _, row in df_req.iterrows():
        line_val = row["LINE_NORM"]
        yr, mo = 2026, 11
        if "WORK_MONTH" in row and pd.notna(row["WORK_MONTH"]):
            wm_str = re.sub(r"\D", "", str(int(float(row["WORK_MONTH"]))))
            if len(wm_str) == 6:
                yr, mo = int(wm_str[:4]), int(wm_str[4:6])

        for day_num, count_col in day_count_cols.items():
            try:
                tgl_obj = date(yr, mo, day_num)
            except ValueError:
                continue

            count_val = parse_count_value(row.get(count_col, 0))
            lead_col = day_lead_cols.get(day_num)
            lead_val = parse_count_value(row.get(lead_col, 0)) if lead_col else 0.0
            records.append({
                "TANGGAL": tgl_obj,
                "LINE": line_val,
                "REQ_SPM_NUM": count_val,
                "LEAD_TIME_NUM": lead_val,
            })

    if not records:
        return empty_df

    df_rec = pd.DataFrame(records)
    # Satu baris Request tersedia untuk setiap kombinasi tanggal + line pada sumber.
    # Count dan Lead Time mengikuti agregasi nilai pada baris Request sumber.
    grouped = df_rec.groupby(["TANGGAL", "LINE"], as_index=False).agg(
        REQ_SPM_NUM=("REQ_SPM_NUM", "sum"),
        LEAD_TIME_NUM=("LEAD_TIME_NUM", "sum"),
    )
    grouped["REQ SPM"] = grouped["REQ_SPM_NUM"].apply(
        lambda value: int(value) if float(value).is_integer() else round(float(value), 2)
    )
    grouped["LEAD TIME"] = grouped["LEAD_TIME_NUM"].apply(
        lambda value: int(value) if float(value).is_integer() else round(float(value), 2)
    )
    return grouped[output_cols]


def _worker_spm_file(spm_folder: str, fname: str, fpath: str, tgl: date | None, use_cache: bool) -> pd.DataFrame:
    if use_cache:
        cached = load_file_cache_if_valid(spm_folder, "SPM", fname, fpath)
        if cached is not None:
            return cached
    df = _parse_spm_excel(fpath, tgl)
    if use_cache:
        save_file_cache(spm_folder, "SPM", fname, fpath, df)
    return df


def _worker_prod_file(prod_folder: str, fname: str, fpath: str, tgl: date | None, use_cache: bool) -> pd.DataFrame:
    if use_cache:
        cached = load_file_cache_if_valid(prod_folder, "PROD", fname, fpath)
        if cached is not None:
            return cached
    df = _parse_packing_excel(fpath, tgl)
    if use_cache:
        save_file_cache(prod_folder, "PROD", fname, fpath, df)
    return df


def _select_excel_files(folder_path: str, date_parser, start_dt: date | None, end_dt: date | None):
    selected_files = []
    legacy_xls = []
    with os.scandir(folder_path) as entries:
        filenames = sorted(entry.name for entry in entries if entry.is_file())

    for fname in filenames:
        if fname.startswith("~$"):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext not in SUPPORTED_EXCEL_EXTENSIONS and ext != ".xls":
            continue
        file_date = date_parser(fname)
        if file_date is not None and not date_in_range(file_date, start_dt, end_dt):
            continue
        full_path = os.path.join(folder_path, fname)
        if ext == ".xls":
            legacy_xls.append(fname)
        else:
            selected_files.append((fname, full_path, file_date))

    if legacy_xls:
        examples = ", ".join(legacy_xls[:3])
        more = " dan file lainnya" if len(legacy_xls) > 3 else ""
        raise ValueError(
            "Ditemukan file Excel format .xls yang tidak dapat dibaca oleh openpyxl: "
            f"{examples}{more}. Simpan ulang file tersebut menjadi .xlsx, lalu proses kembali."
        )
    return selected_files


def load_spm_range_data(
    spm_folder: str,
    start_dt: date | None,
    end_dt: date | None,
    use_cache: bool = True,
    workers: int = 4,
    progress_cb=None,
) -> pd.DataFrame:
    if not spm_folder or not os.path.isdir(spm_folder):
        raise FileNotFoundError(f"Folder SPM tidak ditemukan:\n{spm_folder}")

    if progress_cb:
        progress_cb("Menyaring file SPM sesuai rentang tanggal...")

    selected_files = _select_excel_files(spm_folder, parse_spm_date, start_dt, end_dt)
    if not selected_files:
        return pd.DataFrame()

    dfs = []
    total = len(selected_files)
    done = 0

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {
            pool.submit(_worker_spm_file, spm_folder, fn, fp, dt, use_cache): fn
            for fn, fp, dt in selected_files
        }
        for fut in as_completed(futures):
            fn = futures[fut]
            done += 1
            if progress_cb:
                progress_cb(f"Proses SPM [{done}/{total}]: {fn}")
            part_df = fut.result()
            if part_df is not None and not part_df.empty:
                dfs.append(part_df)

    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)


def load_packing_range_data(
    prod_folder: str,
    start_dt: date | None,
    end_dt: date | None,
    use_cache: bool = True,
    workers: int = 4,
    progress_cb=None,
) -> pd.DataFrame:
    if not prod_folder or not os.path.isdir(prod_folder):
        raise FileNotFoundError(f"Folder Hasil Produksi tidak ditemukan:\n{prod_folder}")

    if progress_cb:
        progress_cb("Menyaring file Hasil Produksi sesuai rentang tanggal...")

    selected_files = _select_excel_files(prod_folder, parse_prod_date, start_dt, end_dt)
    if not selected_files:
        return pd.DataFrame(columns=["TANGGAL", "LINE", "STYLE_KEY", "HASIL PRODUKSI"])

    dfs = []
    total = len(selected_files)
    done = 0

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {
            pool.submit(_worker_prod_file, prod_folder, fn, fp, dt, use_cache): fn
            for fn, fp, dt in selected_files
        }
        for fut in as_completed(futures):
            fn = futures[fut]
            done += 1
            if progress_cb:
                progress_cb(f"Proses Produksi [{done}/{total}]: {fn}")
            part_df = fut.result()
            if part_df is not None and not part_df.empty:
                dfs.append(part_df)

    if not dfs:
        return pd.DataFrame(columns=["TANGGAL", "LINE", "STYLE_KEY", "HASIL PRODUKSI"])
    return pd.concat(dfs, ignore_index=True)


def filter_and_merge_data(
    df_spm_raw: pd.DataFrame,
    df_prod_raw: pd.DataFrame,
    df_req_spm: pd.DataFrame,
    target_lines: set[int] | None,
    style_filters: list[str],
    color_filters: list[str],
    target_status: str,
) -> pd.DataFrame:
    if df_spm_raw is None or df_spm_raw.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    df_spm = df_spm_raw
    if target_lines is not None:
        df_spm = df_spm[df_spm["LINE"].isin(target_lines)]

    status_kw = (target_status or "").strip().upper()
    if status_kw:
        status_values = df_spm["STATUS"].fillna("").astype(str).str.strip().str.upper()
        df_spm = df_spm[status_values.eq(status_kw)]

    if style_filters:
        style_values = df_spm["STYLE"].fillna("").astype(str).str.upper()
        style_mask = pd.Series(False, index=df_spm.index)
        for keyword in style_filters:
            keyword = str(keyword).strip().upper()
            if keyword:
                style_mask |= style_values.str.contains(keyword, regex=False, na=False)
        df_spm = df_spm[style_mask]

    if color_filters:
        color_values = df_spm["COLOR"].fillna("").astype(str).str.upper()
        color_mask = pd.Series(False, index=df_spm.index)
        for keyword in color_filters:
            keyword = str(keyword).strip().upper()
            if keyword:
                color_mask |= color_values.str.contains(keyword, regex=False, na=False)
        df_spm = df_spm[color_mask]

    if df_spm.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    join_keys = ["TANGGAL", "LINE", "STYLE_KEY"]
    df_prod = df_prod_raw
    if df_prod is not None and not df_prod.empty:
        if target_lines is not None:
            df_prod = df_prod[df_prod["LINE"].isin(target_lines)]

        needed_keys = df_spm[join_keys].drop_duplicates()
        df_prod = df_prod.merge(needed_keys, on=join_keys, how="inner", sort=False)
        if not df_prod.empty:
            df_prod = (
                df_prod.groupby(join_keys, as_index=False, sort=False)["HASIL PRODUKSI"]
                .sum()
            )

    if df_prod is None or df_prod.empty:
        df_merged = df_spm.copy()
        df_merged["HASIL PRODUKSI"] = 0
    else:
        df_merged = df_spm.merge(df_prod, how="left", on=join_keys, sort=False)
        df_merged["HASIL PRODUKSI"] = df_merged["HASIL PRODUKSI"].fillna(0)

    df_merged["TANGGAL"] = df_merged["TANGGAL"].apply(_normalize_date_value)
    df_merged["LINE"] = df_merged["LINE"].apply(_normalize_line_value)

    if df_req_spm is not None and not df_req_spm.empty:
        df_req = df_req_spm.copy()
        df_req["TANGGAL"] = df_req["TANGGAL"].apply(_normalize_date_value)
        df_req["LINE"] = df_req["LINE"].apply(_normalize_line_value)
        # Count dan Lead Time dari Andon dapat berupa teks; konversi sebelum merge.
        for metric_col in ("REQ SPM", "LEAD TIME"):
            df_req[metric_col] = df_req[metric_col].apply(parse_count_value)
        if target_lines is not None:
            df_req = df_req[df_req["LINE"].isin(target_lines)]
        df_merged = df_merged.merge(df_req, how="left", on=["TANGGAL", "LINE"], sort=False)
        for metric_col in ("REQ SPM", "LEAD TIME"):
            df_merged[metric_col] = pd.to_numeric(df_merged[metric_col], errors="coerce").fillna(0)
    else:
        df_merged["REQ SPM"] = 0
        df_merged["LEAD TIME"] = 0

    df_merged = df_merged.sort_values(
        ["TANGGAL", "LINE", "STYLE", "COLOR"], kind="stable"
    )
    repeated_style_key = df_merged.duplicated(join_keys, keep="first")
    df_merged.loc[repeated_style_key, "HASIL PRODUKSI"] = pd.NA

    repeated_date_line = df_merged.duplicated(["TANGGAL", "LINE"], keep="first")
    df_merged.loc[repeated_date_line, ["REQ SPM", "LEAD TIME"]] = pd.NA

    return df_merged.loc[:, OUTPUT_COLUMNS].reset_index(drop=True)


def save_styled_excel(df: pd.DataFrame, out_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TRACKING SPM & PRODUKSI"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "landscape"
    ws.print_title_rows = "1:2"

    header_fill = PatternFill(start_color="17365D", end_color="17365D", fill_type="solid")
    header_font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Aptos", size=10, color="243247")
    zebra_fill = PatternFill(start_color="F3F7FC", end_color="F3F7FC", fill_type="solid")
    total_fill = PatternFill(start_color="DCEAF7", end_color="DCEAF7", fill_type="solid")
    total_font = Font(name="Aptos", size=10, bold=True, color="17365D")
    thin = Side(border_style="hair", color="D9E2EF")
    border = Border(bottom=thin)

    req_col_idx = OUTPUT_COLUMNS.index("REQ SPM") + 1
    lead_col_idx = OUTPUT_COLUMNS.index("LEAD TIME") + 1
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
        display_name = "COUNT" if col_name == "REQ SPM" else col_name
        cell = ws.cell(row=2, column=col_idx, value=display_name)
        if col_name == "HASIL PRODUKSI":
            cell.comment = Comment(
                "Hasil produksi pada sumber tersedia di level TANGGAL + LINE + STYLE, bukan per COLOR.",
                "Tracking SPM & Produksi",
            )
        elif col_name == "REQ SPM":
            cell.comment = Comment(
                "REQ SPM (COUNT) diambil dari TROUBLE_NM=Request pada Andon, level TANGGAL + LINE.",
                "Tracking SPM & Produksi",
            )
        elif col_name == "LEAD TIME":
            cell.comment = Comment(
                "Lead Time diambil dari kolom Lead Time Andon pada TROUBLE_NM=Request, level TANGGAL + LINE. Satuan mengikuti sumber Andon.",
                "Tracking SPM & Produksi",
            )
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Header bertingkat: REQ SPM menjadi judul induk untuk COUNT dan LEAD TIME.
    ws.merge_cells(
        start_row=1, start_column=req_col_idx,
        end_row=1, end_column=lead_col_idx,
    )
    parent_header = ws.cell(row=1, column=req_col_idx, value="REQ SPM")
    parent_header.fill = header_fill
    parent_header.font = header_font
    parent_header.alignment = Alignment(horizontal="center", vertical="center")
    parent_header.border = border
    for col_idx in range(req_col_idx, lead_col_idx + 1):
        ws.cell(row=1, column=col_idx).fill = header_fill
        ws.cell(row=1, column=col_idx).border = border

    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 27

    # LINE ditampilkan sebagai bilangan bulat; kuantitas tetap dapat memiliki desimal asli.
    numeric_cols = {"LINE", "ISI BUNDLE", "BUNDLING OUT", "REQ SPM", "LEAD TIME", "HASIL PRODUKSI"}
    left_cols = {"KT ORDER", "STYLE", "COLOR", "DESCRIPTION"}
    widths = {name: max(12, len(name) + 3) for name in OUTPUT_COLUMNS}

    for row_idx, values in enumerate(df.loc[:, OUTPUT_COLUMNS].itertuples(index=False, name=None), start=3):
        zebra = row_idx % 2 == 0
        for col_idx, (col_name, value) in enumerate(zip(OUTPUT_COLUMNS, values), start=1):
            if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)):
                cell_value = None
            else:
                cell_value = value

            cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
            cell.font = data_font
            cell.border = border
            if zebra:
                cell.fill = zebra_fill

            if col_name == "TANGGAL":
                cell.number_format = "dd-mmm-yyyy"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name in numeric_cols:
                # Paksa isi sel menjadi angka Excel, termasuk bila DataFrame berisi string angka.
                if cell_value is not None:
                    numeric_value = parse_count_value(cell_value)
                    is_integer_value = float(numeric_value).is_integer()
                    cell.value = int(numeric_value) if is_integer_value else float(numeric_value)
                    # Hindari pemisah desimal menggantung pada angka bulat (contoh: 125.).
                    # Angka pecahan tetap menampilkan bagian desimalnya.
                    cell.number_format = "#,##0" if col_name == "LINE" or is_integer_value else "#,##0.##"
                else:
                    cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in left_cols:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if cell_value is not None:
                text_len = len(str(cell_value))
                if text_len + 2 > widths[col_name]:
                    widths[col_name] = min(text_len + 2, 42)

    last_data_row = len(df) + 2
    total_row = last_data_row + 1
    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(row=total_row, column=col_idx)
        cell.fill = total_fill
        cell.font = total_font
        cell.border = border
        if col_idx == 1:
            cell.value = "TOTAL"
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_name in {"BUNDLING OUT", "REQ SPM", "LEAD TIME", "HASIL PRODUKSI"}:
            col_letter = get_column_letter(col_idx)
            cell.value = f"=SUM({col_letter}3:{col_letter}{last_data_row})"
            total_value = pd.to_numeric(df[col_name], errors="coerce").sum()
            cell.number_format = "#,##0" if float(total_value).is_integer() else "#,##0.##"
            cell.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[total_row].height = 22

    for col_idx, col_name in enumerate(OUTPUT_COLUMNS, start=1):
        col_letter = get_column_letter(col_idx)
        default_width = {
            "TANGGAL": 15, "LINE": 9, "KT ORDER": 15, "STYLE": 19,
            "COLOR": 22, "DESCRIPTION": 32, "ISI BUNDLE": 14,
            "STATUS": 10, "BUNDLING OUT": 17, "REQ SPM": 13, "LEAD TIME": 13, "HASIL PRODUKSI": 20,
        }.get(col_name, 14)
        ws.column_dimensions[col_letter].width = min(max(widths[col_name], default_width), 42)

    if last_data_row >= 1:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(OUTPUT_COLUMNS))}{last_data_row}"

    wb.save(out_path)


class FastTrackingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SPM & Production Tracking Studio")
        self.geometry("1420x840")
        self.minsize(1120, 700)
        self.configure(bg="#F3F6FB")

        self.df_result = pd.DataFrame(columns=OUTPUT_COLUMNS)
        self.is_running = False
        self.is_exporting = False
        self.ui_queue = queue.Queue()
        self._tree_after_id = None
        self._tree_refresh_token = 0

        self._setup_styles()
        self._build_ui()
        self.after(100, self._poll_ui_queue)

    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        bg_app = "#F3F6FB"
        bg_card = "#FFFFFF"
        primary = "#17365D"
        accent = "#0F9D92"
        text_main = "#172B4D"
        text_muted = "#66758A"
        border = "#DCE4EF"

        style.configure("App.TFrame", background=bg_app)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        style.configure("Card.TLabelframe", background=bg_card, bordercolor=border, relief="solid", borderwidth=1)
        style.configure(
            "Card.TLabelframe.Label", background=bg_card, foreground=primary,
            font=("Segoe UI", 10, "bold"), padding=(4, 0),
        )
        style.configure("CardLabel.TLabel", background=bg_card, foreground=text_main, font=("Segoe UI", 9, "bold"))
        style.configure("Hint.TLabel", background=bg_card, foreground=text_muted, font=("Segoe UI", 8))
        style.configure("Status.TLabel", background=bg_app, foreground=text_muted, font=("Segoe UI", 9))
        style.configure("Section.TLabel", background=bg_card, foreground=primary, font=("Segoe UI", 10, "bold"))
        style.configure("TableHint.TLabel", background=bg_card, foreground=text_muted, font=("Segoe UI", 8))

        style.configure("TEntry", padding=(8, 6), fieldbackground="#FFFFFF", foreground=text_main)
        style.configure("TCombobox", padding=(7, 5), fieldbackground="#FFFFFF", foreground=text_main)
        style.configure("TButton", padding=(10, 6), font=("Segoe UI", 9, "bold"))
        style.configure(
            "Primary.TButton", background=primary, foreground="#FFFFFF", bordercolor=primary,
            padding=(16, 8), font=("Segoe UI", 9, "bold"),
        )
        style.map("Primary.TButton", background=[("active", "#244A78"), ("disabled", "#AAB6C5")])
        style.configure(
            "Secondary.TButton", background="#E8EEF6", foreground=text_main,
            bordercolor=border, padding=(12, 7), font=("Segoe UI", 9),
        )
        style.map("Secondary.TButton", background=[("active", "#D9E3EF")])
        style.configure(
            "Accent.TButton", background=accent, foreground="#FFFFFF", bordercolor=accent,
            padding=(14, 8), font=("Segoe UI", 9, "bold"),
        )
        style.map("Accent.TButton", background=[("active", "#0B8178"), ("disabled", "#AAB6C5")])

        style.configure(
            "Treeview", background="#FFFFFF", foreground=text_main, fieldbackground="#FFFFFF",
            rowheight=29, font=("Segoe UI", 9), bordercolor=border, borderwidth=0,
        )
        style.configure(
            "Treeview.Heading", background="#EAF0F7", foreground=primary,
            font=("Segoe UI", 9, "bold"), relief="flat", padding=(7, 8),
        )
        style.map("Treeview.Heading", background=[("active", "#DCE6F2")])
        style.map("Treeview", background=[("selected", "#D8E9FA")], foreground=[("selected", "#17365D")])

        style.configure("KpiCard.TFrame", background="#FFFFFF", relief="solid", borderwidth=1)
        style.configure("KpiTitle.TLabel", background="#FFFFFF", foreground="#718096", font=("Segoe UI", 8, "bold"))
        style.configure("KpiValue.TLabel", background="#FFFFFF", foreground=primary, font=("Segoe UI", 16, "bold"))

    def _build_ui(self):
        banner = tk.Frame(self, bg="#17365D", padx=22, pady=15)
        banner.pack(fill="x")
        tk.Label(
            banner, text="SPM & PRODUCTION TRACKING", bg="#17365D", fg="#FFFFFF",
            font=("Segoe UI", 17, "bold"),
        ).pack(anchor="w")
        tk.Label(
            banner,
            text="Monitoring bundling, REQ SPM (Andon), dan hasil packing dengan filter fleksibel",
            bg="#17365D", fg="#D7E5F5", font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(3, 0))
        tk.Frame(self, bg="#0F9D92", height=4).pack(fill="x")

        main_wrap = ttk.Frame(self, style="App.TFrame", padding=(18, 14, 18, 14))
        main_wrap.pack(fill="both", expand=True)

        card_src = ttk.LabelFrame(main_wrap, text="  SUMBER DATA  ", style="Card.TLabelframe", padding=(14, 10))
        card_src.pack(fill="x", pady=(0, 10))

        self.var_spm = tk.StringVar(value=DEFAULT_SPM_FOLDER)
        self.var_prod = tk.StringVar(value=DEFAULT_PROD_FOLDER)
        self.var_andon = tk.StringVar(value=DEFAULT_ANDON_FILE)
        self.var_use_cache = tk.BooleanVar(value=True)
        self.var_workers = tk.StringVar(value="4")

        ttk.Label(card_src, text="Folder Data SPM", style="CardLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(card_src, textvariable=self.var_spm).grid(row=0, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(card_src, text="Pilih Folder", style="Secondary.TButton", command=lambda: self._pick_dir(self.var_spm)).grid(row=0, column=2, padx=(6, 0), pady=4)

        ttk.Label(card_src, text="Folder Hasil Produksi", style="CardLabel.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(card_src, textvariable=self.var_prod).grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(card_src, text="Pilih Folder", style="Secondary.TButton", command=lambda: self._pick_dir(self.var_prod)).grid(row=1, column=2, padx=(6, 0), pady=4)

        ttk.Label(card_src, text="File Andon (REQ SPM)", style="CardLabel.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(card_src, textvariable=self.var_andon).grid(row=2, column=1, sticky="ew", padx=4, pady=4)
        ttk.Button(card_src, text="Pilih File", style="Secondary.TButton", command=self._pick_andon_file).grid(row=2, column=2, padx=(6, 0), pady=4)

        source_opts = ttk.Frame(card_src, style="Card.TFrame")
        source_opts.grid(row=3, column=1, columnspan=2, sticky="w", padx=4, pady=(6, 0))
        ttk.Checkbutton(source_opts, text="Gunakan cache lokal per-file", variable=self.var_use_cache).pack(side="left", padx=(0, 20))
        ttk.Label(source_opts, text="Worker paralel", style="CardLabel.TLabel").pack(side="left", padx=(0, 7))
        ttk.Combobox(source_opts, textvariable=self.var_workers, values=["2", "4", "6", "8"], width=4, state="readonly").pack(side="left", padx=(0, 18))
        ttk.Button(source_opts, text="Bersihkan Cache", style="Secondary.TButton", command=self._clear_cache_action).pack(side="left")
        card_src.columnconfigure(1, weight=1)

        card_flt = ttk.LabelFrame(main_wrap, text="  FILTER PENCARIAN  ", style="Card.TLabelframe", padding=(14, 10))
        card_flt.pack(fill="x", pady=(0, 10))

        self.var_tgl_dari = tk.StringVar(value="03-11-2026")
        self.var_tgl_sampai = tk.StringVar(value="26-11-2026")
        self.var_style = tk.StringVar(value="")
        self.var_line = tk.StringVar(value="")
        self.var_color = tk.StringVar(value="")
        self.var_status = tk.StringVar(value="J")

        filter_fields = [
            ("Dari Tanggal", self.var_tgl_dari, 14),
            ("Sampai Tanggal", self.var_tgl_sampai, 14),
            ("Status SPM (J = Jalan)", self.var_status, 8),
            ("Style (opsional)", self.var_style, 22),
            ("Line (opsional)", self.var_line, 14),
            ("Color (opsional)", self.var_color, 22),
        ]
        for index, (label, variable, width) in enumerate(filter_fields):
            row = 0 if index < 3 else 1
            col = (index % 3) * 2
            ttk.Label(card_flt, text=label, style="CardLabel.TLabel").grid(row=row, column=col, sticky="w", padx=(0, 7), pady=6)
            ttk.Entry(card_flt, textvariable=variable, width=width).grid(row=row, column=col + 1, sticky="ew", padx=(0, 18), pady=6)
        for col in (1, 3, 5):
            card_flt.columnconfigure(col, weight=1)

        action_row = ttk.Frame(main_wrap, style="App.TFrame")
        action_row.pack(fill="x", pady=(0, 10))
        self.btn_proses = ttk.Button(action_row, text="▶  PROSES DATA", style="Primary.TButton", command=self.start_tracking_thread)
        self.btn_proses.pack(side="left", padx=(0, 8))
        self.btn_reset = ttk.Button(action_row, text="Reset Filter", style="Secondary.TButton", command=self.reset_filters)
        self.btn_reset.pack(side="left", padx=(0, 8))
        self.btn_export = ttk.Button(action_row, text="Export Excel (.xlsx)", style="Accent.TButton", command=self.export_excel, state="disabled")
        self.btn_export.pack(side="left", padx=(0, 16))
        self.progress = ttk.Progressbar(action_row, mode="indeterminate", length=170)
        self.progress.pack(side="left", padx=(0, 10))
        self.lbl_status = ttk.Label(action_row, text="Siap. Atur filter lalu klik Proses Data.", style="Status.TLabel")
        self.lbl_status.pack(side="left", fill="x", expand=True)

        frm_kpi = ttk.Frame(main_wrap, style="App.TFrame")
        frm_kpi.pack(fill="x", pady=(0, 10))
        for col in range(4):
            frm_kpi.columnconfigure(col, weight=1)
        self.kpi_rows = self._make_kpi_card(frm_kpi, 0, "TOTAL BARIS DATA", "0", padx=(0, 6))
        self.kpi_bundle = self._make_kpi_card(frm_kpi, 1, "TOTAL BUNDLING OUT", "0", padx=(6, 6))
        self.kpi_req_spm = self._make_kpi_card(frm_kpi, 2, "TOTAL REQ SPM", "0", padx=(6, 6))
        self.kpi_prod = self._make_kpi_card(frm_kpi, 3, "TOTAL HASIL PRODUKSI", "0", padx=(6, 0))

        card_tbl = ttk.Frame(main_wrap, style="Card.TFrame", padding=(10, 9))
        card_tbl.pack(fill="both", expand=True)
        table_head = ttk.Frame(card_tbl, style="Card.TFrame")
        table_head.pack(fill="x", pady=(0, 7))
        ttk.Label(table_head, text="HASIL TRACKING", style="Section.TLabel").pack(side="left")
        self.lbl_preview = ttk.Label(table_head, text="Belum ada data", style="TableHint.TLabel")
        self.lbl_preview.pack(side="right")

        table_body = ttk.Frame(card_tbl, style="Card.TFrame")
        table_body.pack(fill="both", expand=True)
        col_widths = {
            "TANGGAL": 112, "LINE": 68, "KT ORDER": 126, "STYLE": 145,
            "COLOR": 170, "DESCRIPTION": 220, "ISI BUNDLE": 104,
            "STATUS": 78, "BUNDLING OUT": 128, "REQ SPM": 105, "LEAD TIME": 110, "HASIL PRODUKSI": 145,
        }
        left_cols = {"KT ORDER", "STYLE", "COLOR", "DESCRIPTION"}
        numeric_cols = {"ISI BUNDLE", "BUNDLING OUT", "REQ SPM", "LEAD TIME", "HASIL PRODUKSI"}

        # Custom two-level table header so REQ SPM can span COUNT + LEAD TIME.
        header_height = 48
        total_header_width = sum(col_widths.get(col, 110) for col in OUTPUT_COLUMNS)
        self.header_canvas = tk.Canvas(
            table_body, height=header_height, highlightthickness=0,
            bd=0, bg="#FFFFFF",
        )
        self.header_canvas.grid(row=0, column=0, sticky="ew")
        header_inner = ttk.Frame(table_body, style="Card.TFrame")
        header_inner.grid_propagate(False)
        self.header_canvas.create_window(
            (0, 0), window=header_inner, anchor="nw",
            width=total_header_width, height=header_height,
        )
        self.header_canvas.configure(scrollregion=(0, 0, total_header_width, header_height))

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS):
            header_inner.columnconfigure(col_idx, minsize=col_widths.get(col_name, 110), weight=0)
        header_inner.rowconfigure(0, minsize=24, weight=1)
        header_inner.rowconfigure(1, minsize=24, weight=1)

        def make_table_header(text):
            return tk.Label(
                header_inner, text=text, bg="#17365D", fg="#FFFFFF",
                font=("Segoe UI", 9, "bold"), relief="solid", bd=1,
                padx=4, pady=3, anchor="center",
            )

        col_idx = 0
        while col_idx < len(OUTPUT_COLUMNS):
            col_name = OUTPUT_COLUMNS[col_idx]
            if col_name == "REQ SPM":
                parent = make_table_header("REQ SPM")
                parent.grid(row=0, column=col_idx, columnspan=2, sticky="nsew")
                make_table_header("COUNT").grid(row=1, column=col_idx, sticky="nsew")
                make_table_header("LEAD TIME").grid(row=1, column=col_idx + 1, sticky="nsew")
                col_idx += 2
                continue

            label = make_table_header(col_name)
            label.grid(row=0, column=col_idx, rowspan=2, sticky="nsew")
            col_idx += 1

        self.tree = ttk.Treeview(table_body, show="", columns=OUTPUT_COLUMNS, selectmode="browse")
        self.tree.tag_configure("odd", background="#F5F8FC")
        self.tree.tag_configure("even", background="#FFFFFF")
        for col in OUTPUT_COLUMNS:
            anchor = "w" if col in left_cols else ("e" if col in numeric_cols else "center")
            self.tree.column(
                col, width=col_widths.get(col, 110), minwidth=65,
                anchor=anchor, stretch=False,
            )

        vsb = ttk.Scrollbar(table_body, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(table_body, orient="horizontal", command=self._scroll_table_x)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=self._sync_table_x)
        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        self.hsb.grid(row=2, column=0, sticky="ew")
        table_body.rowconfigure(1, weight=1)
        table_body.columnconfigure(0, weight=1)

    def _make_kpi_card(self, parent, col: int, title: str, initial_val: str, padx):
        box = ttk.Frame(parent, style="KpiCard.TFrame", padding=(14, 10))
        box.grid(row=0, column=col, sticky="ew", padx=padx)
        ttk.Label(box, text=title, style="KpiTitle.TLabel").pack(anchor="w")
        value_label = ttk.Label(box, text=initial_val, style="KpiValue.TLabel")
        value_label.pack(anchor="w", pady=(3, 0))
        return value_label

    def _scroll_table_x(self, *args):
        """Scroll table rows and the custom two-level header horizontally together."""
        self.tree.xview(*args)
        self.header_canvas.xview(*args)

    def _sync_table_x(self, first, last):
        """Keep the header aligned with the Treeview horizontal scroll position."""
        self.hsb.set(first, last)
        self.header_canvas.xview_moveto(float(first))

    def _pick_dir(self, var):
        initial = var.get().strip()
        selected = filedialog.askdirectory(initialdir=initial if os.path.isdir(initial) else os.getcwd())
        if selected:
            var.set(selected)

    def _pick_andon_file(self):
        initial = os.path.dirname(self.var_andon.get().strip()) or os.getcwd()
        selected = filedialog.askopenfilename(
            initialdir=initial if os.path.isdir(initial) else os.getcwd(),
            filetypes=[("Excel Files", "*.xls *.xlsx *.xlsm"), ("All Files", "*.*")],
        )
        if selected:
            self.var_andon.set(selected)

    def _clear_cache_action(self):
        if self.is_running or self.is_exporting:
            return
        removed = clear_all_cache()
        messagebox.showinfo("Cache Dibersihkan", f"{removed} file cache lokal berhasil dihapus.")

    def reset_filters(self):
        if self.is_running or self.is_exporting:
            return
        self.var_tgl_dari.set("")
        self.var_tgl_sampai.set("")
        self.var_style.set("")
        self.var_line.set("")
        self.var_color.set("")
        self.var_status.set("J")

    def _set_status(self, text: str):
        self.ui_queue.put(("status", text))

    def _poll_ui_queue(self):
        try:
            while True:
                message_type, payload = self.ui_queue.get_nowait()
                if message_type == "status":
                    self.lbl_status.config(text=payload)
                elif message_type == "process_success":
                    self._on_worker_success(payload)
                elif message_type == "process_error":
                    self._on_worker_error(payload)
                elif message_type == "export_success":
                    self._on_export_success(payload)
                elif message_type == "export_error":
                    self._on_export_error(payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_ui_queue)

    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.btn_proses.config(state=state)
        self.btn_reset.config(state=state)
        self.btn_export.config(state="normal" if (not busy and not self.df_result.empty) else "disabled")

    def start_tracking_thread(self):
        if self.is_running or self.is_exporting:
            return
        try:
            start_dt = parse_input_date(self.var_tgl_dari.get())
            end_dt = parse_input_date(self.var_tgl_sampai.get())
            if start_dt and end_dt and start_dt > end_dt:
                raise ValueError("'Dari Tanggal' tidak boleh lebih besar dari 'Sampai Tanggal'.")

            target_lines = parse_line_filter(self.var_line.get())
            style_filters = parse_style_filters(self.var_style.get())
            color_filters = parse_style_filters(self.var_color.get())
            status_kw = self.var_status.get().strip()
            spm_dir = self.var_spm.get().strip()
            prod_dir = self.var_prod.get().strip()
            andon_path = self.var_andon.get().strip()
            use_cache = bool(self.var_use_cache.get())
            workers = int(self.var_workers.get() or "4")
        except Exception as exc:
            messagebox.showerror("Input Filter Tidak Valid", str(exc))
            return

        self.is_running = True
        self._set_busy(True)
        self.progress.start(12)
        self.lbl_status.config(text="Menyiapkan pembacaan data...")
        worker = threading.Thread(
            target=self._worker_run,
            args=(spm_dir, prod_dir, andon_path, target_lines, style_filters, color_filters, status_kw, start_dt, end_dt, use_cache, workers),
            daemon=True,
            name="tracking-data-worker",
        )
        worker.start()

    def _worker_run(self, spm_dir, prod_dir, andon_path, target_lines, style_filters, color_filters, status_kw, start_dt, end_dt, use_cache, workers):
        try:
            df_spm_raw = load_spm_range_data(
                spm_dir, start_dt=start_dt, end_dt=end_dt, use_cache=use_cache,
                workers=workers, progress_cb=self._set_status,
            )
            df_prod_raw = load_packing_range_data(
                prod_dir, start_dt=start_dt, end_dt=end_dt, use_cache=use_cache,
                workers=workers, progress_cb=self._set_status,
            )
            self._set_status("Membaca data REQ SPM (sewing to supermarket)...")
            df_req_spm = load_req_spm_from_andon(andon_path, target_section="sewing to supermarket")

            self._set_status("Memfilter dan menggabungkan data...")
            result_df = filter_and_merge_data(
                df_spm_raw=df_spm_raw, df_prod_raw=df_prod_raw, df_req_spm=df_req_spm,
                target_lines=target_lines, style_filters=style_filters,
                color_filters=color_filters, target_status=status_kw,
            )
            self.ui_queue.put(("process_success", result_df))
        except Exception as exc:
            self.ui_queue.put(("process_error", exc))

    def _on_worker_success(self, result_df: pd.DataFrame):
        self.is_running = False
        self.progress.stop()
        self.df_result = result_df
        self._set_busy(False)
        self._refresh_tree(self.df_result)

        if self.df_result.empty:
            self.lbl_status.config(text="Selesai — tidak ada data yang cocok dengan filter.")
            self.kpi_rows.config(text="0")
            self.kpi_bundle.config(text="0")
            self.kpi_req_spm.config(text="0")
            self.kpi_prod.config(text="0")
        else:
            total_bundle = self.df_result["BUNDLING OUT"].sum()
            total_req_spm = self.df_result["REQ SPM"].sum()
            total_production = self.df_result["HASIL PRODUKSI"].sum()
            self.lbl_status.config(text=f"Selesai — {len(self.df_result):,} baris data berhasil diproses.")
            self.kpi_rows.config(text=f"{len(self.df_result):,}")
            self.kpi_bundle.config(text=f"{total_bundle:,.0f}")
            self.kpi_req_spm.config(text=f"{total_req_spm:,.0f}")
            self.kpi_prod.config(text=f"{total_production:,.0f}")

    def _on_worker_error(self, exc: Exception):
        self.is_running = False
        self.progress.stop()
        self._set_busy(False)
        self.lbl_status.config(text="Proses gagal. Periksa folder atau format file sumber.")
        messagebox.showerror("Error Proses", str(exc))

    @staticmethod
    def _format_tree_value(column: str, value):
        if value is None or (not isinstance(value, (list, tuple, dict)) and pd.isna(value)) or value == "":
            return ""
        if column == "TANGGAL" and hasattr(value, "strftime"):
            return value.strftime("%Y-%m-%d")
        if column in {"ISI BUNDLE", "BUNDLING OUT", "REQ SPM", "LEAD TIME", "HASIL PRODUKSI"}:
            try:
                return f"{float(value):,.0f}"
            except (TypeError, ValueError):
                return value
        return value

    def _refresh_tree(self, df: pd.DataFrame):
        if self._tree_after_id is not None:
            try:
                self.after_cancel(self._tree_after_id)
            except tk.TclError:
                pass
            self._tree_after_id = None

        self._tree_refresh_token += 1
        token = self._tree_refresh_token
        self.tree.delete(*self.tree.get_children())

        if df.empty:
            self.lbl_preview.config(text="0 baris ditampilkan")
            return

        total_rows = len(df)
        preview_rows = min(total_rows, MAX_PREVIEW_ROWS)
        self.lbl_preview.config(
            text=(f"Menyiapkan {preview_rows:,} dari {total_rows:,} baris" if total_rows > preview_rows
                  else f"{total_rows:,} baris")
        )
        rows = df.loc[:, OUTPUT_COLUMNS].itertuples(index=False, name=None)
        self._insert_tree_batch(rows, token, preview_rows, total_rows, 0)

    def _insert_tree_batch(self, rows, token: int, preview_rows: int, total_rows: int, inserted: int):
        if token != self._tree_refresh_token:
            return
        end_at = min(inserted + TREE_BATCH_SIZE, preview_rows)
        for idx in range(inserted, end_at):
            try:
                row_values = next(rows)
            except StopIteration:
                break
            display_values = tuple(
                self._format_tree_value(column, value)
                for column, value in zip(OUTPUT_COLUMNS, row_values)
            )
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert("", "end", values=display_values, tags=(tag,))

        if end_at < preview_rows:
            self.lbl_preview.config(text=f"Menampilkan {end_at:,} dari {total_rows:,} baris...")
            self._tree_after_id = self.after(
                1, lambda: self._insert_tree_batch(rows, token, preview_rows, total_rows, end_at)
            )
        else:
            self._tree_after_id = None
            suffix = f" — preview {preview_rows:,} baris; export tetap seluruh data" if total_rows > preview_rows else ""
            self.lbl_preview.config(text=f"{preview_rows:,} dari {total_rows:,} baris{suffix}")

    def export_excel(self):
        if self.df_result.empty:
            messagebox.showwarning("Data Kosong", "Belum ada data untuk di-export. Silakan klik 'Proses Data' terlebih dahulu.")
            return
        if self.is_running or self.is_exporting:
            return

        style_tag = re.sub(r"[^A-Za-z0-9_-]+", "_", self.var_style.get().strip()) or "ALL_STYLE"
        line_tag = re.sub(r"[^A-Za-z0-9_-]+", "_", self.var_line.get().strip()) or "ALL_LINE"
        color_tag = re.sub(r"[^A-Za-z0-9_-]+", "_", self.var_color.get().strip()) or "ALL_COLOR"
        default_name = f"Tracking_SPM_Produksi_{style_tag}_{line_tag}_{color_tag}.xlsx"
        out_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Workbook", "*.xlsx")], initialfile=default_name,
        )
        if not out_path:
            return

        self.is_exporting = True
        self._set_busy(True)
        self.lbl_status.config(text="Mengekspor file Excel di latar belakang...")
        export_df = self.df_result
        worker = threading.Thread(target=self._worker_export, args=(export_df, out_path), daemon=True, name="tracking-export-worker")
        worker.start()

    def _worker_export(self, df: pd.DataFrame, out_path: str):
        try:
            save_styled_excel(df, out_path)
            self.ui_queue.put(("export_success", out_path))
        except Exception as exc:
            self.ui_queue.put(("export_error", exc))

    def _on_export_success(self, out_path: str):
        self.is_exporting = False
        self._set_busy(False)
        self.lbl_status.config(text="Export Excel selesai.")
        messagebox.showinfo("Export Berhasil", f"File Excel berhasil disimpan ke:\n{out_path}")

    def _on_export_error(self, exc: Exception):
        self.is_exporting = False
        self._set_busy(False)
        self.lbl_status.config(text="Export Excel gagal.")
        messagebox.showerror("Gagal Export", str(exc))


if __name__ == "__main__":
    app = FastTrackingApp()
    app.mainloop()
