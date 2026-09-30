#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
USDA 大豆数据季节性折线图生成脚本
=====================================

功能：
  1. 从 USDA 大豆数据 Excel 文件中读取播种进度、出苗进度、优良率三个指标。
  2. 为每个州（19 个州 + US 全美）生成 3 张季节性折线图，共 57 张单图。
  3. 2026 年数据以红色实线 + 白色填充圆形标记突出显示，并标注百分比数值。
  4. 按指标生成 3 张拼接大图（每张 19 个子图，每排 4 个，US 排首位），
     高分辨率输出（DPI=200），含中文总标题。
  5. 每张图的右下角标注数据来源与数据截至日期（自动取数据中最大周结束日期）。

依赖安装:
  pip install pandas matplotlib openpyxl

数据源 Excel 结构:
  文件默认路径: E:/precip-excel/USDA_Soybean_Data_2020_2026.xlsx

  Sheet 名称（3 个）:
    - Planting_Progress   (播种进度)
    - Emergence_Rate      (出苗进度)
    - Good_Excellent_Rate (优良率)

  列结构（每个 Sheet 一致）:
    | Year | Week_Ending | State | Value |
    |------|-------------|-------|-------|
    | 年份 | 周结束日期    | 州名  | 百分比 |

  州列表（19 个）:
    AR, IA, IL, IN, KS, KY, LA, MI, MN, MO, MS, NC, ND, NE, OH, SD, TN, US, WI

运行方式:
  python generate_usda_charts.py

输出:
  output/
  ├── Planting_Progress/          ← 19 张播种进度单图
  ├── Emergence_Rate/             ← 19 张出苗进度单图
  ├── Good_Excellent_Rate/        ← 19 张优良率单图
  ├── Combined_Planting_Progress.png   ← 播种进度拼接大图
  ├── Combined_Emergence_Rate.png      ← 出苗进度拼接大图
  ├── Combined_Good_Excellent_Rate.png ← 优良率拼接大图
  └── generate_usda_charts.py          ← 本脚本
"""

import os
import sys
import pandas as pd
import matplotlib
matplotlib.use('Agg')  # 非交互式后端，必须置于 import pyplot 之前
import matplotlib.pyplot as plt

# =========================== 配置区（用户可修改） ===========================

# 输入 Excel 文件路径
INPUT_FILE = "./USDA_Soybean_Data_2020_2026.xlsx"

# 输出目录（脚本所在目录）
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# --------------------------- 指标 / Sheet 映射 ---------------------------
SHEET_METRIC_MAP = {
    'Planting_Progress':    'Planting Progress',
    'Emergence_Rate':       'Emergence Rate',
    'Good_Excellent_Rate':  'Good Excellent Rate',
    'Blooming_Rate':        'Blooming Rate',
    'Setting_Pods_Rate':    'Setting Pods Rate',
    'Dropping_Leaves_Rate': 'Dropping Leaves Rate',
    'Harvested_Rate':       'Harvested Rate',
}

# 数据来源标识（显示在每张图右下角）
DATA_SOURCE = 'USDA NASS'

# 拼接大图中文总标题
COMBINED_TITLES = {
    'Planting_Progress':    '美国大豆播种进度（2020-2026）',
    'Emergence_Rate':       '美国大豆出苗进度（2020-2026）',
    'Good_Excellent_Rate':  '美国大豆优良率（2020-2026）',
    'Blooming_Rate':        '美国大豆开花率（2020-2026）',
    'Setting_Pods_Rate':    '美国大豆结荚率（2020-2026）',
    'Dropping_Leaves_Rate': '美国大豆成熟率（2020-2026）',
    'Harvested_Rate':       '美国大豆收割率（2020-2026）',
}

# 年份颜色方案（2020-2025）
YEAR_COLORS = {
    2020: '#2166AC',  # 蓝色
    2021: '#F4A582',  # 橙色
    2022: '#4DAF4A',  # 绿色
    2023: '#762A83',  # 紫色
    2024: '#A6611A',  # 棕色
    2025: '#80CDC1',  # 青色
    # 2026 固定为红色，不在此处定义
}

# =========================== 字体配置 ===========================
plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False


# =========================== 核心函数 ===========================

def build_week_index(df: pd.DataFrame) -> tuple:
    """
    从 DataFrame 构建统一的 X 轴周标签体系和位置映射。

    参数:
        df: 含 Week_Label 列（'%m-%d' 格式字符串）的 DataFrame。

    返回:
        (all_weeks, week_to_pos)
        - all_weeks: 排序后的周标签列表
        - week_to_pos: {周标签: 数值位置} 映射字典
    """
    all_weeks = sorted(
        df['Week_Label'].unique(),
        key=lambda x: (int(x[:2]), int(x[3:]))
    )
    week_to_pos = {wk: i for i, wk in enumerate(all_weeks)}
    return all_weeks, week_to_pos


def plot_state_chart(ax, pivot, all_weeks, week_to_pos, state_name: str,
                     metric_label: str, is_combined: bool = False,
                     avg_pivot=None):
    """
    在给定的 Axes 上绘制单个州的折线图。

    参数:
        ax:          matplotlib Axes 对象
        pivot:       pivot_table 结果，index=Week_Label, columns=Year
        all_weeks:   全局周标签列表（用于 reindex）
        week_to_pos: 周标签到数值 X 坐标的映射
        state_name:  州名
        metric_label: 指标名称（如 "Planting Progress"）
        is_combined: 是否为拼接大图中的子图（影响字号和标签大小）
        avg_pivot:   五年均值 pivot_table（index=Week_Label, columns=Year），可选
    """
    pivot = pivot.reindex(all_weeks)
    years_present = sorted([yr for yr in pivot.columns if pivot[yr].notna().any()])

    # 根据是否拼接模式调整参数
    label_fontsize = 6.5 if is_combined else 8.5
    marker_size = 5 if is_combined else 7
    line_width_2026 = 2.0
    line_width_other = 1.2 if is_combined else 1.5
    legend_fontsize = 6 if is_combined else 9
    title_fontsize = 10 if is_combined else 14
    ylabel = '%' if is_combined else 'Percentage (%)'
    ylabel_fontsize = 8 if is_combined else 11
    xtick_fontsize = 7.5 if is_combined else 9.5
    ytick_fontsize = 7.5 if is_combined else 9.5
    xtick_step = max(1, len(all_weeks) // 12) if is_combined else max(1, len(all_weeks) // 18)
    xytext_offset = 5 if is_combined else 7

    for yr in years_present:
        series = pivot[yr].dropna()
        if series.empty:
            continue
        x_pos = [week_to_pos[wk] for wk in series.index]

        if yr == 2026:
            ax.plot(x_pos, series.values,
                    color='red', linestyle='-', marker='o',
                    markerfacecolor='white', markeredgecolor='red',
                    markersize=marker_size, linewidth=line_width_2026,
                    label=str(yr), zorder=5)
            # 数据标签
            for xi, yi in zip(x_pos, series.values):
                ax.annotate(f'{yi:.0f}', (xi, yi),
                            textcoords="offset points", xytext=(0, xytext_offset),
                            ha='center', va='bottom', fontsize=label_fontsize + 2,
                            color='black', fontweight='normal',
                            clip_on=False)

        else:
            ax.plot(x_pos, series.values,
                    color=YEAR_COLORS.get(yr, '#999999'),
                    linestyle='-', linewidth=line_width_other, label=str(yr))

    # 五年均值线（亮黄色虚线，仅取最近一年，缺失周自然连线）
    if avg_pivot is not None and not avg_pivot.empty:
        max_yr = max(avg_pivot.columns)
        avg_series = avg_pivot[max_yr].dropna()
        if not avg_series.empty:
            x_pos = [week_to_pos[wk] for wk in avg_series.index if wk in week_to_pos]
            y_vals = avg_series.values
            ax.plot(x_pos, y_vals,
                    color='#FFD700', linestyle='--', linewidth=line_width_other + 0.5,
                    label='5Y Avg', zorder=3)

    # 标题、轴标签、图例
    if is_combined:
        ax.set_title(f'{state_name}', fontsize=title_fontsize, fontweight='bold')
    else:
        ax.set_title(f'{state_name} - {metric_label} (2020-2026)',
                     fontsize=title_fontsize, fontweight='bold')

    ax.set_ylabel(ylabel, fontsize=ylabel_fontsize)
    if not is_combined:
        ax.set_xlabel('Week (Month-Day)', fontsize=11)

    ncol = 2 if is_combined else 1
    ax.legend(loc='lower right', fontsize=legend_fontsize, ncol=ncol)
    ax.grid(True, alpha=0.2 if is_combined else 0.25, linestyle='--')

    # Y 轴范围
    y_top = 110 if is_combined else 105
    ax.set_ylim(-5 if is_combined else -2, y_top)

    # X 轴刻度
    tick_idx = list(range(0, len(all_weeks), xtick_step))
    ax.set_xticks(tick_idx)
    ax.set_xticklabels([all_weeks[i] for i in tick_idx],
                       rotation=45, ha='right', fontsize=xtick_fontsize)
    ax.tick_params(axis='y', labelsize=ytick_fontsize)


def build_avg_pivot(state_df, df_5y, state):
    """
    构建五年均值 pivot。优先用 USDA API 数据，缺失时本地计算 2022-2026 均值。
    返回 (avg_pivot, source) — source 为 "api" 或 "local"。
    """
    # 尝试 API 数据
    if df_5y is not None and not df_5y.empty:
        state_5y = df_5y[df_5y['State'] == state].copy()
        if not state_5y.empty:
            # 只取最近一年（如 2025）的五年均值
            max_yr = state_5y['Year'].max()
            state_5y = state_5y[state_5y['Year'] == max_yr]
            avg_pivot = state_5y.pivot_table(
                index='Week_Label', columns='Year', values='Value', aggfunc='first'
            )
            return avg_pivot, "api"

    # 降级：本地计算 2022-2026 五年均值（按周聚合）
    state_5y_local = state_df[state_df['Year'].between(2022, 2026)]
    if not state_5y_local.empty:
        avg_series = state_5y_local.groupby('Week_Label')['Value'].mean().round(1)
        avg_pivot = pd.DataFrame({2025: avg_series})  # 列名用 2025
        return avg_pivot, "local"

    return None, "none"


def add_source_footer(fig, data_as_of: str, data_source: str = DATA_SOURCE,
                      is_combined: bool = False):
    """
    在图形右下角添加数据来源与数据截至日期标注。

    参数:
        fig:         matplotlib Figure 对象
        data_as_of:  数据截至日期（字符串，如 '2026-09-25'）
        data_source: 数据来源名称
        is_combined: 是否为拼接大图（影响字号）
    """
    fontsize = 11 if is_combined else 7.5
    fig.text(0.995, 0.005,
             f'数据来源: {data_source} | 数据截至: {data_as_of}',
             ha='right', va='bottom', fontsize=fontsize,
             color='#666666', zorder=1000)


def generate_single_charts(df, all_weeks, week_to_pos, sheet_name, metric_label, df_5y=None,
                           data_source=DATA_SOURCE, data_as_of=''):
    """
    生成单个指标的所有州折线图（19 张），保存到 output/<sheet_name>/ 目录。
    """
    sub_dir = os.path.join(OUTPUT_DIR, sheet_name)
    os.makedirs(sub_dir, exist_ok=True)

    states_all = sorted(df['State'].unique())
    ordered_states = ['US'] + sorted([s for s in states_all if s != 'US'])

    for state in ordered_states:
        state_df = df[df['State'] == state].copy()
        pivot = state_df.pivot_table(
            index='Week_Label', columns='Year', values='Value', aggfunc='first'
        )

        # 五年均值 pivot（API 优先，缺失时本地计算）
        avg_pivot, _ = build_avg_pivot(state_df, df_5y, state)

        fig, ax = plt.subplots(figsize=(14, 7))
        plot_state_chart(ax, pivot, all_weeks, week_to_pos, state, metric_label,
                         is_combined=False, avg_pivot=avg_pivot)
        plt.tight_layout()

        # 右下角标注数据来源与数据截至日期
        add_source_footer(fig, data_as_of, data_source)

        safe_state = state.replace(' ', '_')
        filepath = os.path.join(sub_dir, f'{safe_state}_{sheet_name}.png')
        fig.savefig(filepath, dpi=150, bbox_inches='tight')
        plt.close(fig)
        print(f"  [单图] {safe_state}_{sheet_name}.png")

    print(f"  → {sheet_name} 单图完成 ({len(ordered_states)} 张)")


def generate_combined_chart(df, all_weeks, week_to_pos, sheet_name, metric_label, df_5y=None,
                            data_source=DATA_SOURCE, data_as_of=''):
    """
    生成单个指标的拼接大图（5 行 × 4 列，US 首位），保存到 output/ 目录。
    """
    states_all = sorted(df['State'].unique())
    ordered_states = ['US'] + sorted([s for s in states_all if s != 'US'])

    n_states = len(ordered_states)
    n_cols = 4
    n_rows = (n_states + n_cols - 1) // n_cols

    # 每个子图 14×7 英寸
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(14 * n_cols, 7 * n_rows))
    axes = axes.flatten()

    for idx, state in enumerate(ordered_states):
        ax = axes[idx]
        state_df = df[df['State'] == state].copy()
        pivot = state_df.pivot_table(
            index='Week_Label', columns='Year', values='Value', aggfunc='first'
        )

        # 五年均值 pivot（API 优先，缺失时本地计算）
        avg_pivot, _ = build_avg_pivot(state_df, df_5y, state)

        plot_state_chart(ax, pivot, all_weeks, week_to_pos, state, metric_label,
                         is_combined=True, avg_pivot=avg_pivot)

    # 隐藏多余子图
    for idx in range(n_states, n_rows * n_cols):
        axes[idx].set_visible(False)

    fig.suptitle(COMBINED_TITLES[sheet_name], fontsize=24, fontweight='bold', y=0.99)
    plt.tight_layout(rect=[0, 0, 1, 0.97])

    # 右下角标注数据来源与数据截至日期
    add_source_footer(fig, data_as_of, data_source, is_combined=True)

    combined_path = os.path.join(OUTPUT_DIR, f'Combined_{sheet_name}.png')
    fig.savefig(combined_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    print(f"  [大图] Combined_{sheet_name}.png")


# =========================== 主流程 ===========================

def main():
    print("=" * 60)
    print("USDA 大豆数据图表生成脚本")
    print(f"输入文件: {INPUT_FILE}")
    print(f"输出目录: {OUTPUT_DIR}")
    print("=" * 60)

    # ---- 检查输入文件 ----
    if not os.path.exists(INPUT_FILE):
        print(f"[错误] 找不到输入文件: {INPUT_FILE}")
        print("请修改脚本顶部的 INPUT_FILE 变量后重试。")
        sys.exit(1)

    # ---- 逐指标处理 ----
    for sheet_name, metric_label in SHEET_METRIC_MAP.items():
        print(f"\n{'─' * 40}")
        print(f"处理指标: {metric_label} (Sheet: {sheet_name})")

        # 读取数据
        df_raw = pd.read_excel(INPUT_FILE, sheet_name=sheet_name)
        df_raw['Week_Ending'] = pd.to_datetime(df_raw['Week_Ending'])
        df_raw['Week_Label'] = df_raw['Week_Ending'].dt.strftime('%m-%d')
        df_raw['Year'] = df_raw['Year'].astype(str)

        # 分离五年均值行（Year 含 _5Y 后缀）
        df_5y = df_raw[df_raw['Year'].str.contains('_5Y')].copy()
        df_5y['Year'] = df_5y['Year'].str.replace('_5Y', '').astype(int)
        df = df_raw[~df_raw['Year'].str.contains('_5Y')].copy()
        df['Year'] = df['Year'].astype(int)

        # 构建全局 X 轴
        all_weeks, week_to_pos = build_week_index(df)
        print(f"  周范围: {all_weeks[0]} ~ {all_weeks[-1]} ({len(all_weeks)} 周)")

        # 数据截至日期 = 数据中最大周结束日期
        data_as_of = df['Week_Ending'].max().strftime('%Y-%m-%d')
        print(f"  数据来源: {DATA_SOURCE} | 数据截至: {data_as_of}")

        # 步骤 1: 生成单图
        print("  生成单图...")
        generate_single_charts(df, all_weeks, week_to_pos, sheet_name, metric_label, df_5y,
                               DATA_SOURCE, data_as_of)

        # 步骤 2: 生成拼接大图
        print("  生成拼接大图...")
        generate_combined_chart(df, all_weeks, week_to_pos, sheet_name, metric_label, df_5y,
                                DATA_SOURCE, data_as_of)

    # ---- 完成 ----
    print(f"\n{'=' * 60}")
    single_count = 19 * len(SHEET_METRIC_MAP)
    combined_count = len(SHEET_METRIC_MAP)
    print(f"全部完成！共生成 {single_count} 张单图 + {combined_count} 张拼接大图。")
    print(f"输出目录: {OUTPUT_DIR}")
    print("=" * 60)


if __name__ == '__main__':
    main()
