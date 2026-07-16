"""
生產日報表彙總
把多個工作表（各站人員填寫的日報表）依「日期＋料號＋站」彙總，
同一天、同一料號、同一站若分散在多列／多個工作表，加總 OK／NG 數量。
"""
import re
from io import BytesIO

import openpyxl
import pandas as pd

_HEADER_ROW = 4
_SUBHEADER_ROW = 5
_DATA_START_ROW = 6
_DATE_COL = 1          # A：日期
_ITEM_COL = 2           # B：料號
_STATION_START_COL = 4  # D 開始才是製程站；C 欄是「發料數量」不算站


def _cell(ws, row, col) -> str:
    v = ws.cell(row=row, column=col).value
    return "" if v is None else str(v).strip()


def _station_map(ws) -> list[dict]:
    """掃描第 4/5 列標題，找出各站的欄位範圍與 OK/NG 欄位。"""
    max_col = ws.max_column
    starts = []
    for c in range(_STATION_START_COL, max_col + 1):
        v = _cell(ws, _HEADER_ROW, c)
        if v:
            starts.append((c, v.split("\n")[0].strip()))

    seen: dict[str, int] = {}
    stations = []
    for i, (c, name) in enumerate(starts):
        end = starts[i + 1][0] - 1 if i + 1 < len(starts) else max_col
        ok_col = ng_col = None
        for cc in range(c, end + 1):
            sub = _cell(ws, _SUBHEADER_ROW, cc)
            if sub == "OK":
                ok_col = cc
            elif sub == "NG":
                ng_col = cc

        # 同一份表若有重複站名（例如流程中出現兩次「外觀檢查」），加序號區分，避免彙總時誤合併
        seen[name] = seen.get(name, 0) + 1
        display_name = name if seen[name] == 1 else f"{name}{seen[name]}"

        stations.append({"name": display_name, "start": c, "ok_col": ok_col, "ng_col": ng_col})
    return stations


def _month_from_sheet(ws) -> int | None:
    m = re.match(r"(\d+)", _cell(ws, 2, 1))
    return int(m.group(1)) if m else None


def _num(ws, row, col) -> float:
    if col is None:
        return 0.0
    v = ws.cell(row=row, column=col).value
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def parse_daily_report_workbook(wb: openpyxl.Workbook) -> pd.DataFrame:
    """
    解析生產日報表活頁簿，回傳長表：欄位為 月, 日, 料號, 站, OK, NG, 工作表。
    自動略過沒有「日期／料號」標題列的工作表（例如空白總表）。
    """
    records = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row < _DATA_START_ROW:
            continue
        if _cell(ws, _HEADER_ROW, _DATE_COL).split("\n")[0].strip() != "日期" or \
           _cell(ws, _HEADER_ROW, _ITEM_COL).split("\n")[0].strip() != "料號":
            continue

        month = _month_from_sheet(ws)
        stations = _station_map(ws)

        for r in range(_DATA_START_ROW, ws.max_row + 1):
            item_no = _cell(ws, r, _ITEM_COL)
            date_raw = _cell(ws, r, _DATE_COL)
            if not item_no or not date_raw:
                continue
            try:
                day = int(float(date_raw))
            except ValueError:
                continue

            for st_info in stations:
                ok = _num(ws, r, st_info["ok_col"])
                ng = _num(ws, r, st_info["ng_col"])
                if st_info["ok_col"] is None and st_info["ng_col"] is None:
                    # 沒有 OK/NG 細分（例如「包裝」），該站起始欄位的值視為數量
                    ok = _num(ws, r, st_info["start"])
                if ok == 0 and ng == 0:
                    continue
                records.append({
                    "月": month,
                    "日": day,
                    "料號": item_no,
                    "站": st_info["name"],
                    "OK": ok,
                    "NG": ng,
                    "工作表": sheet_name,
                })

    return pd.DataFrame(records, columns=["月", "日", "料號", "站", "OK", "NG", "工作表"])


def aggregate_daily_report(df: pd.DataFrame) -> pd.DataFrame:
    """
    依「日期＋料號＋站」彙總加總 OK／NG，回傳寬表：
    每列為一個 (月, 日, 料號)，各站的 OK/NG 各佔一欄。
    """
    if df.empty:
        return pd.DataFrame()

    grouped = df.groupby(["月", "日", "料號", "站"], as_index=False)[["OK", "NG"]].sum()
    station_order = list(dict.fromkeys(df["站"]))

    wide = grouped.pivot(index=["月", "日", "料號"], columns="站", values=["OK", "NG"]).fillna(0)
    ordered_cols = [(kind, station) for station in station_order for kind in ("OK", "NG")]
    wide = wide.reindex(columns=pd.MultiIndex.from_tuples(ordered_cols))
    wide.columns = [f"{station}_{kind}" for kind, station in wide.columns]
    wide = wide.reset_index().sort_values(["月", "日", "料號"]).reset_index(drop=True)
    for col in wide.columns[3:]:
        wide[col] = wide[col].astype(int)
    return wide


def generate_report_excel(wide_df: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wide_df.to_excel(writer, index=False, sheet_name="報表彙總")
    output.seek(0)
    return output
