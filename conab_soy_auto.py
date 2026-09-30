# -*- coding: utf-8 -*-
"""
CONAB 巴西大豆种植率 / 收获率 —— 自动抓取 + Excel 输出 + 多年对比图
==================================================================
数据源：CONAB "Progresso de Safra"（周报 Acompanhamento das Lavouras / Plantio e Colheita）
官网（2025 起迁至 gov.br 域名）：
    https://www.gov.br/conab/pt-br/atuacao/informacoes-agropecuarias/safras/progresso-de-safra

功能：
    1) 自动翻页收集全部周报链接（b_start 分页）
    2) 每期定位 plantio-e-colheita 入口，直链下载 xlsx（PK 头）
    3) 解析 Soja 块：Semeadura -> planting（种植率），Colheita -> harvest（收获率），
       取"本周"列（日期行最后一个日期对应的值）
    4) 与已有历史 CSV 增量合并（去重）
    5) 输出 Excel（长表 + 种植率透视 + 收获率透视 + meta）
    6) 生成多年分对比图：每个州 + 全国("12 estados") 的种植率、收获率各一张，
       x 轴按"产季对齐"（产季自当年 9 月 1 日起，单位=周）
    7) 每张图右下角标注数据来源与数据截至日期（自动取数据中最大日期）

依赖：requests / openpyxl / python-calamine / pandas / matplotlib
运行：
    python conab_soy_auto.py            # 增量模式（读已有CSV + 抓官网新增）
    python conab_soy_auto.py --full     # 全量模式（从官网全部重抓）
"""

import argparse
import io
import json
import os
import re
import sys
import time
import warnings
from datetime import date, datetime, timedelta

import pandas as pd

import requests

warnings.filterwarnings("ignore")

# ----------------------------------------------------------------------------
# 常量与配置
# ----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LIST_URL = ("https://www.gov.br/conab/pt-br/atuacao/informacoes-"
            "agropecuarias/safras/progresso-de-safra")
CACHE_DIR = os.path.join(BASE_DIR, "_cache_xlsx")
OUT_CSV = os.path.join(BASE_DIR, "CONAB_soy_progress_auto.csv")
OUT_XLSX = os.path.join(BASE_DIR, "CONAB_soy_progress.xlsx")
OUT_META = os.path.join(BASE_DIR, "CONAB_soy_progress_meta.json")
CHART_DIR = os.path.join(BASE_DIR, "charts")

# 数据来源标识（显示在每张图右下角）
DATA_SOURCE = "CONAB Progresso de Safra"

# 历史基线：此前已交付的合并数据（2022-04 ~ 2026-06），作为增量合并的底
HISTORY_CSVS = [
    os.path.join(BASE_DIR, "CONAB_soy_progress_2022_2026.csv"),
]

# 官方年度归档 zip（2022-2024，旧版数据，用于全量重建历史底）
ARCHIVE_TOKENS = {
    2022: "lvdzd4M6",
    2023: "oTQZ6N06",
    2024: "__RgwV3s",
}
ARCHIVE_URL = ("https://arquivosportal.conab.gov.br/api/public/dl/{token}/"
               "progressodesafra/{year}.zip")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# 巴西大豆主产州（CONAB "12 estados" 口径），保持输出顺序
STATES = [
    "Tocantins", "Maranhão", "Piauí", "Bahia", "Mato Grosso",
    "Mato Grosso do Sul", "Goiás", "Minas Gerais", "São Paulo",
    "Paraná", "Santa Catarina", "Rio Grande do Sul",
]
TOTAL_STATE = "12 estados"

# 产季：种植自 9 月启动；产季标签 = "起始年/后两位"
SEASON_START_MONTH = 9  # 9 月 1 日

# ----------------------------------------------------------------------------
# 手动补丁：2024/25 年度 12 estados（全国）收获率（百分比）
# 用于补齐官网周报缺失的 2025-01~03 段（每周一个值，自 2025-01-04 起）。
# 已有日期（2025-04-05 起）也会被覆盖为同值，结果一致、无副作用。
# ----------------------------------------------------------------------------
MANUAL_PATCH_HARVEST_2024_25_12ESTADOS = [
    ("2025-01-04", 0.0), ("2025-01-11", 0.3), ("2025-01-18", 1.2),
    ("2025-01-25", 3.2), ("2025-02-01", 8.0), ("2025-02-08", 14.8),
    ("2025-02-15", 25.5), ("2025-02-22", 36.4), ("2025-03-01", 48.4),
    ("2025-03-08", 60.9), ("2025-03-15", 69.8), ("2025-03-22", 76.4),
    ("2025-03-29", 81.4), ("2025-04-05", 85.3), ("2025-04-12", 88.3),
    ("2025-04-19", 92.5), ("2025-04-26", 94.8), ("2025-05-03", 97.7),
    ("2025-05-10", 98.5), ("2025-05-17", 98.9), ("2025-05-24", 99.5),
    ("2025-05-31", 99.8),
]

# ----------------------------------------------------------------------------
# 手动补丁：2021/22 年度 12 estados（全国）收割进度（百分比）
# 用于补齐官网周报缺失的 2022-01~02 段（每周一个值，自 2022-01-15 起）。
# 已有日期（2022-04-23 起）会被覆盖为同值，结果一致、无副作用。
# ----------------------------------------------------------------------------
MANUAL_PATCH_HARVEST_2021_22_12ESTADOS = [
    ("2022-01-15", 1.7), ("2022-01-22", 5.5), ("2022-01-29", 12.3),
    ("2022-02-05", 16.8), ("2022-02-12", 25.0), ("2022-02-19", 33.0),
    ("2022-02-26", 42.1), ("2022-03-05", 52.2), ("2022-03-12", 63.1),
    ("2022-03-19", 70.6), ("2022-03-26", 75.8), ("2022-04-02", 81.2),
    ("2022-04-09", 84.9), ("2022-04-16", 87.1), ("2022-04-23", 90.8),
    ("2022-04-30", 93.9), ("2022-05-07", 94.9), ("2022-05-14", 96.8),
    ("2022-05-21", 98.1), ("2022-05-28", 99.1), ("2022-06-04", 99.4),
    ("2022-06-11", 99.8),
]


def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Referer": LIST_URL,
        "Accept": "*/*",
    })
    return s


# ----------------------------------------------------------------------------
# 1. 收集周报链接（翻页）
# ----------------------------------------------------------------------------
def get_weekly_links(session: requests.Session) -> dict:
    """返回 {url: 'Acompanhamento das Lavouras - dd/mm a dd/mm/yy'}，按日期排序"""
    links = {}
    b = 0
    while True:
        url = LIST_URL if b == 0 else f"{LIST_URL}?b_start:int={b}"
        try:
            r = session.get(url, timeout=30)
            r.raise_for_status()
        except Exception as e:  # noqa: BLE001
            print(f"[warn] 列表页失败 b={b}: {e}")
            break
        text = r.text
        found = 0
        for href, label in re.findall(r'<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>',
                                      text, re.S):
            t = re.sub(r"<[^>]+>", "", label).strip()
            t = re.sub(r"\s+", " ", t)
            if re.match(r"Acompanhamento das Lavouras - \d{2}/\d{2} a \d{2}/\d{2}/\d{2}", t):
                if href not in links:
                    links[href] = t
                    found += 1
        if not found:
            break
        b += 20
        if b > 500:
            break
    items = sorted(links.items(), key=lambda kv: _label_date(kv[1]))
    return dict(items)


def _label_date(label: str) -> tuple:
    """返回 (年, 月, 日)，取周报结束日期（'a dd/mm/yy'）"""
    m = re.search(r"(\d{2})/(\d{2}) a (\d{2})/(\d{2})/(\d{2})$", label)
    if not m:
        return (0, 0, 0)
    yy = 2000 + int(m.group(5))
    mm = int(m.group(4))  # 结束月
    dd = int(m.group(3))  # 结束日
    return (yy, mm, dd)


# ----------------------------------------------------------------------------
# 2. 定位 plantio-e-colheita 入口并下载 xlsx
# ----------------------------------------------------------------------------
def get_plantio_entry(session: requests.Session, detail_url: str):
    try:
        r = session.get(detail_url, timeout=30)
        r.raise_for_status()
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 详情页失败 {detail_url}: {e}")
        return None
    m = re.search(r'href="([^"]*plantio-e-colheita[^"]*)"', r.text)
    return m.group(1) if m else None


def download_xlsx(session: requests.Session, entry_url: str) -> bytes | None:
    try:
        r = session.get(entry_url, timeout=90)
        r.raise_for_status()
        data = r.content
        if not data.startswith(b"PK"):
            print(f"[warn] 非 xlsx 响应: {entry_url} ({len(data)}B)")
            return None
        return data
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 下载失败 {entry_url}: {e}")
        return None


def fetch_weekly_xlsx(session: requests.Session, detail_url: str) -> bytes | None:
    entry = get_plantio_entry(session, detail_url)
    if not entry:
        return None
    return download_xlsx(session, entry)


# ----------------------------------------------------------------------------
# 3. 解析 Soja 块（通用：新版多日期列 / 旧版均兼容）
# ----------------------------------------------------------------------------
def _to_rows(xlsx_bytes: bytes) -> list[list]:
    """优先 openpyxl，报错回退 calamine，统一返回行列表（空值->None）"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
        return rows
    except TypeError:
        pass
    except Exception:  # noqa: BLE001
        pass
    try:
        import calamine  # 旧版本模块名
    except ImportError:
        import python_calamine as calamine  # 0.8.x 新模块名
    wb = calamine.CalamineWorkbook.from_filelike(io.BytesIO(xlsx_bytes))
    sheet = wb.get_sheet_by_name(wb.sheet_names[0])
    return sheet.to_python()


def _parse_date(cell) -> date | None:
    if isinstance(cell, datetime):
        return cell.date()
    if isinstance(cell, date):
        return cell
    if isinstance(cell, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(cell.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _norm_state(name: str) -> str:
    name = str(name).replace("nan", "").strip()
    if name.lower().startswith("12 estado"):
        return TOTAL_STATE
    return name


def parse_soja_xlsx(xlsx_bytes: bytes):
    """返回 list[(date, metric, state, value)]，metric in {'planting','harvest'}"""
    rows = _to_rows(xlsx_bytes)
    out = []
    i = 0
    n = len(rows)
    while i < n:
        cells = rows[i]
        joined = " ".join(str(c) for c in cells if c is not None)
        # 定位 Soja 块
        if "soja" in joined.lower() and ("safra" in joined.lower() or "estados" in joined.lower()):
            # 在 Soja 块内找 Semeadura / Colheita 指标行
            j = i + 1
            while j < n:
                jcells = rows[j]
                jjoined = " ".join(str(c) for c in jcells if c is not None)
                # 遇到下一个作物块（非空标题且含 Safra，且不含 estados 说明行）则退出 Soja 块
                if (j > i + 1 and jjoined.strip()
                        and re.search(r" - Safra \d", jjoined)
                        and "estados" not in jjoined.lower()
                        and "semeadura" not in jjoined.lower()
                        and "colheita" not in jjoined.lower()):
                    break
                if re.search(r"(?i)\bSemeadura\b", jjoined):
                    out.extend(_parse_metric_block(rows, j, "planting"))
                elif re.search(r"(?i)\bColheita\b", jjoined):
                    out.extend(_parse_metric_block(rows, j, "harvest"))
                j += 1
            i = j
        else:
            i += 1
    return out


def _parse_metric_block(rows: list, idx: int, metric: str):
    """从指标行(如 'Semeadura')开始，向下找表头/日期行，再取数据行"""
    n = len(rows)
    # 找表头行：含 'Estado' 或 'Unidade da Feder'
    hdr_i = None
    j = idx + 1
    while j < n:
        joined = " ".join(str(c) for c in rows[j] if c is not None)
        if ("estado" in joined.lower() or "unidade da feder" in joined.lower()) and \
           ("semana" in joined.lower() or "média" in joined.lower() or "media" in joined.lower()):
            hdr_i = j
            break
        # 安全阀：遇到另一个指标/块
        if joined.strip() and re.search(r"(?i)(semeadura|colheita)", joined):
            return []
        j += 1
    if hdr_i is None:
        return []
    # 找日期行：表头后 1~3 行内，含 >=2 个可解析日期（3 列日期 -> 本周=最后一个）
    dates_row_i = None
    for k in range(hdr_i + 1, min(hdr_i + 4, n)):
        ds = [_parse_date(c) for c in rows[k]]
        ds = [d for d in ds if d]
        if len(ds) >= 2:
            dates_row_i = k
            break
    if dates_row_i is None:
        return []
    # 本周列索引 = 日期行最后一个日期的列
    last_dt = None
    last_col = None
    for ci, c in enumerate(rows[dates_row_i]):
        d = _parse_date(c)
        if d is not None:
            last_dt = d
            last_col = ci
    if last_dt is None:
        return []
    # 数据行
    out = []
    k = dates_row_i + 1
    while k < n:
        cells = rows[k]
        joined = " ".join(str(c) for c in cells if c is not None)
        if not joined.strip():
            break
        # 若进入新块标题则停
        if re.search(r" - Safra \d", joined) and "estados" not in joined.lower():
            break
        st_name = cells[1] if len(cells) > 1 else None
        st = _norm_state(st_name)
        if st and last_col < len(cells):
            val = cells[last_col]
            fv = _to_float(val)
            if fv is not None:
                out.append((last_dt, metric, st, fv))
        k += 1
    return out


def _to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        v = v.strip().replace(",", ".")
        if v in ("", "-", "—"):
            return None
        try:
            return float(v)
        except ValueError:
            return None
    return None


# ----------------------------------------------------------------------------
# 4. 官方归档 zip（全量重建历史底）
# ----------------------------------------------------------------------------
def fetch_archive(year: int, session: requests.Session):
    """下载并解析指定年份官方归档，返回 DataFrame 行(list of dict)"""
    token = ARCHIVE_TOKENS.get(year)
    if not token:
        return []
    url = ARCHIVE_URL.format(token=token, year=year)
    try:
        r = session.get(url, timeout=180)
        r.raise_for_status()
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 归档 {year} 下载失败: {e}")
        return []
    import zipfile
    try:
        z = zipfile.ZipFile(io.BytesIO(r.content))
        names = [x for x in z.namelist()
                 if x.lower().endswith((".xlsx", ".xls")) and "plantio" in x.lower()]
        if not names:
            names = [x for x in z.namelist() if x.lower().endswith((".xlsx", ".xls"))]
        recs = []
        for nm in names[:5]:
            data = z.read(nm)
            recs.extend(parse_soja_xlsx(data))
        return recs
    except Exception as e:  # noqa: BLE001
        print(f"[warn] 归档 {year} 解析失败: {e}")
        return []


# ----------------------------------------------------------------------------
# 5. 主流程
# ----------------------------------------------------------------------------
def load_history() -> pd.DataFrame | None:
    for p in HISTORY_CSVS:
        if os.path.exists(p):
            df = pd.read_csv(p)
            df["date"] = pd.to_datetime(df["date"])
            return df
    return None


def apply_manual_patch(df: pd.DataFrame) -> pd.DataFrame:
    """合入 12 estados 收割进度手动补丁（必须在去重前调用）。

    支持多个产季补丁（2024/25、2021/22 等），补丁值以百分比记录，
    写入时除以 100 转成 0-1 比例；与已有行同 (date, metric, state)
    时以补丁为准（keep='first'），因此既有数据被同值覆盖、结果一致。
    """
    patches = [
        MANUAL_PATCH_HARVEST_2024_25_12ESTADOS,
        MANUAL_PATCH_HARVEST_2021_22_12ESTADOS,
    ]
    rows = []
    for data in patches:
        rows += [{"date": d, "metric": "harvest", "state": "12 estados",
                  "value": pct / 100.0, "week_end": d}
                 for d, pct in data]
    patch = pd.DataFrame(rows, columns=["date", "metric", "state", "value", "week_end"])
    patch["date"] = pd.to_datetime(patch["date"])
    patch["week_end"] = pd.to_datetime(patch["week_end"])
    print(f"[info] 合入手动补丁 {len(patch)} 行（12 estados 收割进度，"
          f"{len(patches)} 个产季）")
    return pd.concat([patch, df], ignore_index=True)


def build_frames(full: bool = False):
    """返回 (long_df, meta)"""
    session = _session()
    meta = {"source": "CONAB Progresso de Safra（官网周报 Plantio e Colheita xlsx）",
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M")}

    # 历史基线
    hist = load_history()
    if hist is not None and not full:
        print(f"[info] 使用历史基线 {len(hist)} 行")
        meta["history_base"] = os.path.basename(HISTORY_CSVS[0])
    else:
        hist = pd.DataFrame(columns=["date", "metric", "state", "value", "week_end"])
        if full:
            # 全量：从官方归档重建 2022-2024 底
            recs = []
            for yr in sorted(ARCHIVE_TOKENS):
                recs += fetch_archive(yr, session)
            if recs:
                hist = pd.DataFrame(recs, columns=["date", "metric", "state", "value"])
                hist["week_end"] = hist["date"]
                hist["date"] = pd.to_datetime(hist["date"])
                print(f"[info] 官方归档重建 {len(hist)} 行")

    # 增量抓取官网周报
    links = get_weekly_links(session)
    print(f"[info] 官网周报共 {len(links)} 期")
    new_rows = []
    for url, label in links.items():
        dt = _label_date(label)
        if not hist.empty:
            # 跳过已覆盖（历史已含该周结束日）
            if (hist["date"].dt.year == dt[0]).any() and \
               ((hist["date"].dt.year == dt[0]) & (hist["date"].dt.month == dt[1])).sum() > 0:
                continue
        xlsx = fetch_weekly_xlsx(session, url)
        if not xlsx:
            continue
        recs = parse_soja_xlsx(xlsx)
        print(f"   {label}: soja {len(recs)} 行")
        new_rows.extend(recs)
        time.sleep(0.3)

    if new_rows:
        newdf = pd.DataFrame(new_rows, columns=["date", "metric", "state", "value"])
        newdf["week_end"] = newdf["date"]
        hist = pd.concat([hist, newdf], ignore_index=True)

    # 合入手动补丁（补齐 2024/25 12 estados 收获率缺失段），随后统一清洗+去重
    hist = apply_manual_patch(hist)

    # 统一清洗
    hist["date"] = pd.to_datetime(hist["date"])
    hist["week_end"] = pd.to_datetime(hist["week_end"])
    # 异常值修正：>1 视为百分比写错（除以 100）
    mask = hist["value"] > 1.0
    if mask.any():
        print(f"[info] 修正 {int(mask.sum())} 个 >1 异常值（百分比/100）")
        hist.loc[mask, "value"] = hist.loc[mask, "value"] / 100.0
    hist["value"] = hist["value"].clip(0.0, 1.0)
    # 去重
    hist = hist.drop_duplicates(subset=["date", "metric", "state"]).sort_values(
        ["date", "metric", "state"]).reset_index(drop=True)
    # 产季列
    hist = add_season_weeks(hist)

    meta["rows"] = int(len(hist))
    meta["coverage"] = (f"{hist['date'].min().date()} 至 {hist['date'].max().date()}"
                        if not hist.empty else "空")
    meta["states"] = sorted(hist["state"].unique().tolist())
    meta["metrics"] = {"planting": "种植率 Semeadura", "harvest": "收获率 Colheita"}
    return hist, meta


def season_of(d: pd.Timestamp) -> str:
    if d.month >= SEASON_START_MONTH:
        y0 = d.year
    else:
        y0 = d.year - 1
    return f"{y0}/{str((y0 + 1) % 100).zfill(2)}"


def season_week(d: pd.Timestamp, metric: str) -> float:
    """产季周序（周）。
    - planting（种植率）：自当年 9 月 1 日（播种季启动）起算
    - harvest（收获率）：自当年 12 月下旬（12 月 20 日，收获季启动）起算
    """
    y0 = d.year if d.month >= SEASON_START_MONTH else d.year - 1
    if metric == "harvest":
        start = pd.Timestamp(year=y0, month=12, day=20)
    else:
        start = pd.Timestamp(year=y0, month=SEASON_START_MONTH, day=1)
    return (d - start).days / 7.0


def add_season_weeks(df: pd.DataFrame) -> pd.DataFrame:
    """统一补充 season / season_week 列（按指标区分周序基准）"""
    df = df.copy()
    df["season"] = df["date"].apply(season_of)
    df["season_week"] = df.apply(lambda r: season_week(r["date"], r["metric"]),
                                 axis=1)
    return df


def clean_season(df: pd.DataFrame) -> pd.DataFrame:
    """
    产季残留清洗：CONAB 周报在产季早期（如 9 月）仍保留上一产季完成的极值
    （如收获率=1.0），随后跳回 0 再上升，会在多年对比图上形成垂直假线。
    规则：对每个 (state, metric, season) 序列，若出现相邻"从≈1 突降到≈0"
    （差值 > 0.5）的跳变，则丢弃该跳变点之前的全部残留点。
    """
    out = []
    for (st, metric, season), g in df.groupby(["state", "metric", "season"]):
        g = g.sort_values("date")
        vals = g["value"].tolist()
        drop_at = None
        for i in range(1, len(vals)):
            if vals[i - 1] - vals[i] > 0.5:
                drop_at = i
                break
        if drop_at:
            g = g.iloc[drop_at:]
        out.append(g)
    if not out:
        return df
    return pd.concat(out, ignore_index=True)


# ----------------------------------------------------------------------------
# 6. Excel 输出
# ----------------------------------------------------------------------------
def _fmt_dt(ser: pd.Series) -> pd.Series:
    """datetime 列转 'YYYY-MM-DD'，否则原样保留（兼容离线模式）"""
    if pd.api.types.is_datetime64_any_dtype(ser):
        return ser.dt.strftime("%Y-%m-%d")
    if ser.map(lambda x: isinstance(x, str)).all() and not ser.empty:
        return ser
    return ser.astype(str)


def export_excel(df: pd.DataFrame, path: str):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        long_tab = df[["date", "week_end", "season", "season_week", "metric",
                       "state", "value"]].copy()
        for col in ("date", "week_end"):
            long_tab[col] = _fmt_dt(long_tab[col])
        long_tab.to_excel(w, sheet_name="明细", index=False)
        for metric, label in [("planting", "种植率"), ("harvest", "收获率")]:
            sub = df[df["metric"] == metric]
            pivot = sub.pivot_table(index="date", columns="state", values="value",
                                    aggfunc="last")
            pivot = pivot[ [c for c in pivot.columns if c in STATES] + 
                          ([TOTAL_STATE] if TOTAL_STATE in pivot.columns else []) ]
            pivot.index = [_fmt_dt(pd.Series([d]))[0] for d in pivot.index]
            pivot.to_excel(w, sheet_name=label + "透视")
    print(f"[ok] Excel 已写出: {path}")


# ----------------------------------------------------------------------------
# 7. 多年对比图
# ----------------------------------------------------------------------------
def plot_comparisons(df: pd.DataFrame, out_dir: str,
                     data_source: str = DATA_SOURCE, data_as_of: str = ''):
    os.makedirs(out_dir, exist_ok=True)
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import PercentFormatter
    plt.rcParams.update({
        "font.size": 9,
        "font.sans-serif": ["Noto Sans CJK SC"],
        "axes.unicode_minus": False,
    })

    states = [c for c in STATES if c in set(df["state"])]
    if TOTAL_STATE in set(df["state"]):
        states = states + [TOTAL_STATE]

    for metric, mlabel, color in [
        ("planting", "种植率", "#1f77b4"),
        ("harvest", "收获率", "#d62728"),
    ]:
        for st in states:
            sub = df[(df["metric"] == metric) & (df["state"] == st)]
            sub = sub[sub["season_week"] <= 25]  # 仅保留产季 0-25 周
            if sub.empty:
                continue
            fig, ax = plt.subplots(figsize=(9, 5.5))
            seasons = sorted(sub["season"].unique())
            latest_season = seasons[-1]  # 最新产季 = season 值最大的一条曲线
            cmap = plt.get_cmap("tab10")
            # 其他年份专用色板：tab10 排除索引3（即红色 #d62728），
            # 避免与最新年度撞色，保证每条线颜色各不相同
            other_cmap = [cmap(i) for i in (0, 1, 2, 4, 5, 6, 7, 8, 9)]
            for si, sname in enumerate(seasons):
                s = sub[sub["season"] == sname].sort_values("date")
                # 最新年度线条用红色，其余年份按顺序循环取色板（无红色重复）
                c = "#d62728" if sname == latest_season else other_cmap[si % 9]
                ax.plot(s["season_week"], s["value"] * 100.0,
                        marker="o", markersize=3, linewidth=1.5,
                        color=c, label=sname)
                # 最新年度在每个数据点上方显示数值标注，其余年份不加
                if sname == latest_season:
                    for x, y in zip(s["season_week"], s["value"] * 100.0):
                        ax.annotate(f"{y:.1f}%", (x, y),
                                    textcoords="offset points", xytext=(0, 6),
                                    ha="center", fontsize=7, color=c)
            if st == TOTAL_STATE:
                title = ("巴西全国收割率进度对比" if metric == "harvest"
                         else "巴西全国种植率进度对比")
            else:
                title = f"{st} · 大豆{mlabel}多年对比（CONAB Progresso de Safra）"
            ax.set_title(title, fontsize=14)  # 标题字号明显加大两号
            ax.set_xlim(0, 25)
            # X 轴：周序 -> 日期刻度（基准日按指标区分：
            # 种植率自 9/1 起算、收获率自 12/20（12月下旬）起算）
            tick_weeks = list(range(0, 26, 2))
            if metric == "harvest":
                ref_start = date(2023, 12, 20)
                xlabel = "日期（收获季自12月下旬起，仅显示 0-25 周）"
            else:
                ref_start = date(2023, 9, 1)
                xlabel = "日期（产季自 9月1日起，仅显示 0-25 周）"
            ax.set_xticks(tick_weeks)
            ax.set_xticklabels([(ref_start + timedelta(weeks=w)).strftime("%m-%d")
                                for w in tick_weeks])
            ax.set_xlabel(xlabel)
            ax.set_ylabel(f"{mlabel} (%)")
            ax.set_ylim(-2, 105)
            ax.yaxis.set_major_formatter(PercentFormatter())
            ax.grid(alpha=0.3, linestyle="--")
            # 图例放图表下方，列数按产季数量自动（最多 6 列）
            n_seasons = len(seasons)
            ax.legend(title="产季", loc="upper center",
                      bbox_to_anchor=(0.5, -0.10),
                      ncol=min(6, max(1, n_seasons)), fontsize=8)
            # 为底部图例留足空间，避免被裁剪
            fig.tight_layout(rect=[0, 0, 1, 1])
            fig.subplots_adjust(bottom=0.18)
            # 右下角标注数据来源与数据截至日期
            fig.text(0.995, 0.005,
                     f"数据来源: {data_source} | 数据截至: {data_as_of}",
                     ha="right", va="bottom", fontsize=7, color="#666666")
            fname = os.path.join(out_dir,
                                 f"{st}_{metric}_by_season.png".replace(" ", "_"))
            fig.savefig(fname, dpi=150)
            plt.close(fig)
            print(f"   [图] {fname}")


# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="CONAB 巴西大豆种植率/收获率 自动抓取")
    ap.add_argument("--full", action="store_true", help="全量模式（重建历史底+全部重抓）")
    ap.add_argument("--no-crawl", action="store_true",
                    help="不联网，仅用历史基线生成 Excel 和图（离线演示）")
    args = ap.parse_args()

    os.makedirs(CACHE_DIR, exist_ok=True)
    os.makedirs(CHART_DIR, exist_ok=True)

    if args.no_crawl:
        hist = load_history()
        if hist is None:
            print("[err] 无历史基线，无法离线运行")
            sys.exit(1)
        # 合入手动补丁（补齐 2024/25 12 estados 收获率缺失段）
        hist = apply_manual_patch(hist)
        # 与 build_frames() 相同的值清洗：>1 视为百分比写错（除以 100），再 clip 到 [0,1]
        mask = hist["value"] > 1.0
        if mask.any():
            print(f"[info] 修正 {int(mask.sum())} 个 >1 异常值（百分比/100）")
            hist.loc[mask, "value"] = hist.loc[mask, "value"] / 100.0
        hist["value"] = hist["value"].clip(0.0, 1.0)
        # 与 build_frames() 相同的去重：补丁在前 keep='first'，补丁值优先
        hist = hist.drop_duplicates(subset=["date", "metric", "state"]).sort_values(
            ["date", "metric", "state"]).reset_index(drop=True)
        meta = {"mode": "offline", "rows": int(len(hist)),
                "coverage": f"{hist['date'].min().date()} 至 {hist['date'].max().date()}"}
        hist = add_season_weeks(hist)
    else:
        hist, meta = build_frames(full=args.full)

    hist.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"[ok] CSV 已写出: {OUT_CSV} ({len(hist)} 行)")
    with open(OUT_META, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    export_excel(hist, OUT_XLSX)
    # 数据截至日期 = 数据中最大日期
    data_as_of = hist["date"].max().strftime("%Y-%m-%d") if not hist.empty else ""
    print(f"数据来源: {DATA_SOURCE} | 数据截至: {data_as_of}")
    plot_comparisons(clean_season(hist), CHART_DIR, data_as_of=data_as_of)

    print("\n=== 汇总 ===")
    print(f"行数: {len(hist)}")
    print(f"覆盖: {meta.get('coverage')}")
    print(f"Excel: {OUT_XLSX}")
    print(f"图表目录: {CHART_DIR}")


if __name__ == "__main__":
    main()
