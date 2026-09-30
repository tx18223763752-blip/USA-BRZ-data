#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CFTC Managed Money 持仓数据下载与可视化（Disaggregated Futures-Only）
=====================================================================
品种：美豆 SOYBEANS / 美豆油 SOYBEAN OIL / 美豆粕 SOYBEAN MEAL / 美玉米 CORN / 美小麦 WHEAT(SRW)，均为 CBOT
字段：money_manager_positions_long_all / short_all / spread_all（Managed Money 分类）
数据源：
  首选 Socrata API：https://publicreporting.cftc.gov/resource/72hh-3qpy.json（免 Key）
  若 API 不可达（如 403），自动回退 CFTC 官方历史压缩包 fut_disagg_txt_YYYY.zip（同一数据集、同字段语义）
逻辑：
  首次运行全量拉取 2015-01-01 至今，落盘固定 CSV；
  后续运行仅按本地最大 report_date 增量追加最新一期（CFTC 每周发布一期）；
  按 (report_date, market) 去重排序，保证幂等。
绘图：
  5 张品种多年份对比图（2020-至今，X 轴年内周序；历史年份灰色细线，最新年份红色折线 + 白填充红框标记点）
  + 1 张 5 品种 2020-至今净持仓（多头-空头）连续趋势图
  = 6 张子图按 2 行 x 3 列拼接为一张大图；每张子图右下角标注数据来源与数据截至日期。

用法：
  python cftc_managed_money.py            # 增量更新（无本地 CSV 时自动全量）
  python cftc_managed_money.py --full     # 强制全量重建
  python cftc_managed_money.py --skip-plot  # 仅更新数据，不绘图
"""

import os
import io
import csv
import json
import time
import zipfile
import argparse
import datetime as dt
import urllib.request

import urllib.parse  # noqa: E402

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# ========= matplotlib中文配置（GitHub Actions 专用）=========
plt.rcParams["font.sans-serif"] = ["Noto Sans CJK SC", "WenQuanYi Zen Hei"]
plt.rcParams["axes.unicode_minus"] = False
# ===========================================================

# ---------------------------------------------------------------------------
# 常量配置
# ---------------------------------------------------------------------------
OUT_DIR = r"E:\precip-excel\CFTC-Managed Money"
CSV_PATH = os.path.join(OUT_DIR, "cftc_managed_money_data.csv")
PNG_PATH = os.path.join(OUT_DIR, "cftc_managed_money_6panel.png")

START_DATE = "2015-01-01"
TODAY = dt.date.today()
CURRENT_YEAR = TODAY.year

SOCRATA_URL = "https://publicreporting.cftc.gov/resource/72hh-3qpy.json"
HISTORY_ZIP_URL = "https://www.cftc.gov/files/dea/history/fut_disagg_txt_{year}.zip"

# 品种 -> 官方市场名（market_and_exchange_names 精确匹配，避免其他交易所同名合约）
MARKETS = {
    "SOYBEANS": "SOYBEANS - CHICAGO BOARD OF TRADE",
    "SOYBEAN OIL": "SOYBEAN OIL - CHICAGO BOARD OF TRADE",
    "SOYBEAN MEAL": "SOYBEAN MEAL - CHICAGO BOARD OF TRADE",
    "CORN": "CORN - CHICAGO BOARD OF TRADE",
    "WHEAT": "WHEAT-SRW - CHICAGO BOARD OF TRADE",
}

# Socrata 字段名
SOC_FIELDS = [
    "report_date",
    "market_and_exchange_names",
    "money_manager_positions_long_all",
    "money_manager_positions_short_all",
    "money_manager_positions_spread_all",
]

# 官方 zip 字段名 -> 统一字段名（与 Socrata 对齐）
ZIP_FIELD_MAP = {
    "Market_and_Exchange_Names": "market_and_exchange_names",
    "Report_Date_as_YYYY-MM-DD": "report_date",
    "M_Money_Positions_Long_All": "money_manager_positions_long_all",
    "M_Money_Positions_Short_All": "money_manager_positions_short_all",
    "M_Money_Positions_Spread_All": "money_manager_positions_spread_all",
}

OUTPUT_COLUMNS = [
    "report_date",
    "market",
    "market_and_exchange_names",
    "money_manager_positions_long_all",
    "money_manager_positions_short_all",
    "money_manager_positions_spread_all",
]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept": "application/json,text/csv,text/plain,*/*"}


# ---------------------------------------------------------------------------
# 网络请求
# ---------------------------------------------------------------------------
def http_get_text(url, timeout=60, retries=2):
    """带 UA 的 GET，返回文本；失败重试。"""
    return http_get_bytes(url, timeout=timeout, retries=retries).decode("utf-8", errors="replace")


def http_get_bytes(url, timeout=60, retries=2):
    """带 UA 的 GET，返回原始字节；失败重试。"""
    last_err = None
    for i in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as e:
            last_err = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"GET {url} 失败: {last_err}")


# ---------------------------------------------------------------------------
# 数据获取：Socrata API（首选）
# ---------------------------------------------------------------------------
def fetch_socrata_since(since_date):
    """从 Socrata API 拉取 report_date >= since_date 的目标品种行。返回 DataFrame。"""
    frames = []
    offset = 0
    limit = 50000
    while True:
        where = f"report_date >= '{since_date}'"
        query = (
            f"$select={','.join(SOC_FIELDS)}"
            f"&$where={urllib.parse.quote(where)}"
            f"&$order=report_date"
            f"&$limit={limit}&$offset={offset}"
        )
        url = f"{SOCRATA_URL}?{query}"
        txt = http_get_text(url, timeout=60)
        rows = json.loads(txt)
        if not isinstance(rows, list) or not rows:
            break
        frames.append(pd.DataFrame(rows))
        if len(rows) < limit:
            break
        offset += limit
    if not frames:
        return pd.DataFrame(columns=SOC_FIELDS + ["market"])
    df = pd.concat(frames, ignore_index=True)
    return normalize_socrata(df)


def normalize_socrata(df):
    df = df.copy()
    df["report_date"] = pd.to_datetime(df["report_date"].astype(str).str[:10])
    for c in ["money_manager_positions_long_all", "money_manager_positions_short_all",
              "money_manager_positions_spread_all"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["market"] = df["market_and_exchange_names"].map({v: k for k, v in MARKETS.items()})
    df = df[df["market"].notna()].copy()
    return df[OUTPUT_COLUMNS]


# ---------------------------------------------------------------------------
# 数据获取：CFTC 官方历史压缩包（回退）
# ---------------------------------------------------------------------------
def fetch_zip_years(years):
    """下载并解析指定年份的官方 zip（fut_disagg_txt_YYYY.zip）。返回 DataFrame。"""
    frames = []
    for y in sorted(set(years)):
        url = HISTORY_ZIP_URL.format(year=y)
        raw_bytes = http_get_bytes(url, timeout=120, retries=3)
        zf = zipfile.ZipFile(io.BytesIO(raw_bytes))
        name = zf.namelist()[0]
        raw = zf.read(name).decode("utf-8", errors="replace")
        reader = csv.DictReader(io.StringIO(raw))
        df = pd.DataFrame(reader)
        df = df.rename(columns=ZIP_FIELD_MAP)
        df = df[list(ZIP_FIELD_MAP.values())]
        df["report_date"] = pd.to_datetime(df["report_date"].astype(str).str[:10])
        for c in ["money_manager_positions_long_all", "money_manager_positions_short_all",
                  "money_manager_positions_spread_all"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        df["market"] = df["market_and_exchange_names"].map({v: k for k, v in MARKETS.items()})
        df = df[df["market"].notna()].copy()
        frames.append(df)
        print(f"  已解析 {y}: {len(df)} 行")
    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    return pd.concat(frames, ignore_index=True)[OUTPUT_COLUMNS]


def fetch_from_zip_since(since_date):
    """从官方 zip 拉取 report_date >= since_date 的目标行。"""
    since = pd.to_datetime(since_date)
    year_floor = since.year
    years = list(range(year_floor, CURRENT_YEAR + 1))
    df = fetch_zip_years(years)
    df = df[df["report_date"] >= since]
    return df


# ---------------------------------------------------------------------------
# 主下载流程
# ---------------------------------------------------------------------------
def load_local_max_date():
    if not os.path.exists(CSV_PATH):
        return None
    df = pd.read_csv(CSV_PATH, dtype={"report_date": str})
    if df.empty:
        return None
    return pd.to_datetime(df["report_date"]).max()


def save_data(df):
    df = df.sort_values(["report_date", "market"]).reset_index(drop=True)
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    print(f"数据已保存: {CSV_PATH}（{len(df)} 行）")


def update_data(full=False):
    """全量/增量更新数据文件。"""
    if full:
        print("== 全量模式：拉取 2015-01-01 至今 ==")
        new_rows = None
        # 首选 Socrata API
        try:
            print("尝试 Socrata API ...")
            new_rows = fetch_socrata_since(START_DATE)
            print(f"Socrata API 成功，获取 {len(new_rows)} 行")
        except Exception as e:
            print(f"Socrata API 不可用（{e}），回退 CFTC 官方历史压缩包 ...")
            new_rows = fetch_from_zip_since(START_DATE)
        save_data(new_rows)
        return new_rows

    max_date = load_local_max_date()
    if max_date is None:
        print("本地无数据，转全量模式 ...")
        return update_data(full=True)

    # 增量：CFTC 每周发布，取本地最大日期之后一天开始
    since = (max_date + pd.Timedelta(days=1)).strftime("%Y-%m-%d")
    print(f"== 增量模式：本地最大 report_date = {max_date.date()}，拉取 >= {since} ==")
    new_rows = None
    try:
        print("尝试 Socrata API ...")
        new_rows = fetch_socrata_since(since)
        print(f"Socrata API 成功，获取 {len(new_rows)} 行")
    except Exception as e:
        print(f"Socrata API 不可用（{e}），回退 CFTC 官方历史压缩包 ...")
        new_rows = fetch_from_zip_since(since)

    old = pd.read_csv(CSV_PATH, dtype={"report_date": str})
    old["report_date"] = pd.to_datetime(old["report_date"])
    combined = pd.concat([old, new_rows], ignore_index=True)
    combined = combined.drop_duplicates(subset=["report_date", "market"], keep="last")
    combined = combined.sort_values(["report_date", "market"]).reset_index(drop=True)
    save_data(combined)
    return combined


# ---------------------------------------------------------------------------
# 绘图
# ---------------------------------------------------------------------------
def setup_font():
    plt.rcParams["axes.unicode_minus"] = False
    for f in ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]:
        try:
            plt.rcParams["font.sans-serif"] = [f, "DejaVu Sans"]
            return
        except Exception:
            continue


def annotate_source(ax, max_date):
    ax.text(0.99, 0.01,
            f"数据来源：CFTC Disaggregated Report (Managed Money)\n数据截至：{max_date:%Y-%m-%d}",
            transform=ax.transAxes, ha="right", va="bottom",
            fontsize=7.5, color="#444444",
            bbox=dict(boxstyle="round,pad=0.35", fc="white", ec="#CCCCCC", alpha=0.85))


def plot_yearly_comparison(ax, df, market, cn_name, max_date_overall):
    """多年份对比：X 轴年内周序，2020-至今各年份不同颜色，最新年份红色折线+白填充红框标记。"""
    sub = df[df["market"] == market].copy()
    sub["year"] = sub["report_date"].dt.year
    # 年内第几周（按年内日序计算）：1月1日为第1周，单调递增到52/53，避免ISO周跨年回绕造成年末→年初的连接线
    sub["week"] = ((sub["report_date"] - pd.to_datetime(sub["year"].astype(str) + "-01-01")).dt.days // 7) + 1
    latest_year = sub["year"].max()

    hist_years = list(range(2020, latest_year))
    if hist_years:
        # 历史年份 6 色（避开红色，红色保留给最新年份），保证每年一色互不重复
        HIST_COLORS = ["#1F77B4", "#FF7F0E", "#2CA02C", "#9467BD", "#8C564B", "#17BECF"]
        colors = {y: HIST_COLORS[i % len(HIST_COLORS)] for i, y in enumerate(hist_years)}
        for y in hist_years:
            yd = sub[sub["year"] == y]
            if yd.empty:
                continue
            ax.plot(yd["week"], yd["net"], color=colors[y], linewidth=1.1,
                    alpha=0.85, zorder=3)

    yd = sub[sub["year"] == latest_year].sort_values("week")
    if not yd.empty:
        # 红色折线 + 白填充红框标记点
        ax.plot(yd["week"], yd["net"], color="#D62728", linewidth=2.0, zorder=5)
        ax.fill_between(yd["week"], yd["net"], 0, color="#D62728", alpha=0.06, zorder=1)
        ax.plot(yd["week"], yd["net"], "o", color="white", markeredgecolor="#D62728",
                markeredgewidth=1.3, markersize=3.6, zorder=6)

    ax.set_title(f"{cn_name}（{market}）净持仓多年对比", fontsize=11)
    ax.set_xlabel("年内周序", fontsize=8.5)
    ax.set_ylabel("净持仓（多头-空头）", fontsize=8.5)
    ax.set_xlim(0, 53)
    ax.grid(True, which="both", axis="y", linestyle="--", alpha=0.35)
    ax.tick_params(labelsize=8)
    # 图例：历史年份各色 + 最新年红色
    handles = []
    for y in hist_years:
        handles.append(plt.Line2D([0], [0], color=colors[y], linewidth=1.1,
                                  label=str(y)))
    if not yd.empty:
        handles.append(plt.Line2D([0], [0], color="#D62728", linewidth=2, marker="o",
                                  markerfacecolor="white", markeredgecolor="#D62728",
                                  markeredgewidth=1.3, markersize=5, label=str(latest_year)))
    ax.legend(handles=handles, loc="upper left", fontsize=7.5, framealpha=0.9,
              ncol=2 if len(handles) > 4 else 1)
    max_date = sub["report_date"].max() if not sub.empty else max_date_overall
    annotate_source(ax, max_date)


def plot_all_trend(ax, df, max_date_overall):
    """5 品种 2020-至今 净持仓连续趋势图。"""
    colors = {"SOYBEANS": "#1F77B4", "SOYBEAN OIL": "#FF7F0E",
              "SOYBEAN MEAL": "#2CA02C", "CORN": "#9467BD", "WHEAT": "#8C564B"}
    sub = df[df["report_date"] >= "2020-01-01"].copy()
    for market, color in colors.items():
        md = sub[sub["market"] == market].sort_values("report_date")
        if md.empty:
            continue
        ax.plot(md["report_date"], md["net"], color=color, linewidth=1.5,
                label={"SOYBEANS": "美豆", "SOYBEAN OIL": "美豆油",
                       "SOYBEAN MEAL": "美豆粕", "CORN": "美玉米",
                       "WHEAT": "美小麦"}[market], zorder=3)
    ax.set_title("五大品种净持仓连续趋势（2020 至今）", fontsize=11)
    ax.set_xlabel("报告日期", fontsize=8.5)
    ax.set_ylabel("净持仓（多头-空头）", fontsize=8.5)
    ax.grid(True, which="both", linestyle="--", alpha=0.35)
    ax.tick_params(labelsize=8)
    ax.legend(loc="upper left", fontsize=8, framealpha=0.9)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    annotate_source(ax, max_date_overall)


def make_plot(df):
    """生成 2x3 拼接大图。"""
    df = df.copy()
    df["net"] = df["money_manager_positions_long_all"] - df["money_manager_positions_short_all"]

    setup_font()
    fig, axes = plt.subplots(2, 3, figsize=(21, 12), dpi=150)
    max_date_overall = df["report_date"].max()

    specs = [
        ("SOYBEANS", "美豆"),
        ("SOYBEAN OIL", "美豆油"),
        ("SOYBEAN MEAL", "美豆粕"),
        ("CORN", "美玉米"),
        ("WHEAT", "美小麦"),
    ]
    for i, (market, cn) in enumerate(specs):
        plot_yearly_comparison(axes[i // 3][i % 3], df, market, cn, max_date_overall)

    plot_all_trend(axes[1][2], df, max_date_overall)

    fig.suptitle("CFTC Managed Money 持仓（Disaggregated Futures-Only, CBOT）", fontsize=15, y=0.99)
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    fig.savefig(PNG_PATH, dpi=150, bbox_inches="tight")
    print(f"拼接大图已生成: {PNG_PATH}")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="CFTC Managed Money 持仓数据下载与可视化")
    parser.add_argument("--full", action="store_true", help="强制全量重建数据")
    parser.add_argument("--skip-plot", action="store_true", help="仅更新数据不绘图")
    args = parser.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)

    df = update_data(full=args.full)
    if df is None or df.empty:
        print("无数据，退出")
        return 1

    if not args.skip_plot:
        make_plot(df)
    print("完成。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
