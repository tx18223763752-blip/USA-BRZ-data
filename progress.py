import requests
import pandas as pd
import time
import pickle
import os
from typing import List, Dict, Optional, Tuple

# ==================== 配置 ====================
API_KEY = "89A74005-9309-3F2A-A510-AB8C0ABC0574"
BASE_URL = "https://quickstats.nass.usda.gov/api/api_GET/"

# 18 个主产州
STATE_CODES = {
    "AR": "Arkansas", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "MI": "Michigan",
    "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri", "NE": "Nebraska",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "SD": "South Dakota", "TN": "Tennessee", "WI": "Wisconsin"
}
STATES = list(STATE_CODES.keys())

REQUEST_TIMEOUT = 30
RETRY_TIMES = 3
REQUEST_DELAY = 2.0

# 缓存与输出路径
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "soybean_cache.pkl")
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "USDA_Soybean_Data_2020_2026.xlsx")

# 所有需要抓取的基础指标定义
# (metric_name, statisticcat_desc, unit_desc)
METRIC_DEFS = [
    # ---- 普通值 ----
    ("Planting_Progress",   "PROGRESS",          "PCT PLANTED"),
    ("Emergence_Rate",      "PROGRESS",          "PCT EMERGED"),
    ("Blooming_Rate",       "PROGRESS",          "PCT BLOOMING"),
    ("Setting_Pods_Rate",   "PROGRESS",          "PCT SETTING PODS"),
    ("Dropping_Leaves_Rate","PROGRESS",          "PCT DROPPING LEAVES"),
    ("Harvested_Rate",      "PROGRESS",          "PCT HARVESTED"),
    ("Good_Rate",           "CONDITION",         "PCT GOOD"),
    ("Excellent_Rate",      "CONDITION",         "PCT EXCELLENT"),
    # ---- 五年均值 ----
    ("Planting_Progress_5Y",   "PROGRESS, 5 YEAR AVG",  "PCT PLANTED"),
    ("Emergence_Rate_5Y",      "PROGRESS, 5 YEAR AVG",  "PCT EMERGED"),
    ("Blooming_Rate_5Y",       "PROGRESS, 5 YEAR AVG",  "PCT BLOOMING"),
    ("Setting_Pods_Rate_5Y",   "PROGRESS, 5 YEAR AVG",  "PCT SETTING PODS"),
    ("Dropping_Leaves_Rate_5Y","PROGRESS, 5 YEAR AVG",  "PCT DROPPING LEAVES"),
    ("Harvested_Rate_5Y",      "PROGRESS, 5 YEAR AVG",  "PCT HARVESTED"),
    ("Good_Rate_5Y",           "CONDITION, 5 YEAR AVG", "PCT GOOD"),
    ("Excellent_Rate_5Y",      "CONDITION, 5 YEAR AVG", "PCT EXCELLENT"),
]

# Sheet 定义: (Sheet名称, 普通值metric_name, 五年均值metric_name, 是否需要合并Good+Excellent)
SHEET_DEFS = [
    ("Planting_Progress",    "Planting_Progress",    "Planting_Progress_5Y",    False),
    ("Emergence_Rate",       "Emergence_Rate",       "Emergence_Rate_5Y",       False),
    ("Good_Excellent_Rate",  "Good_Rate",            "Good_Rate_5Y",            True),
    ("Blooming_Rate",        "Blooming_Rate",        "Blooming_Rate_5Y",        False),
    ("Setting_Pods_Rate",    "Setting_Pods_Rate",    "Setting_Pods_Rate_5Y",    False),
    ("Dropping_Leaves_Rate", "Dropping_Leaves_Rate", "Dropping_Leaves_Rate_5Y", False),
    ("Harvested_Rate",       "Harvested_Rate",       "Harvested_Rate_5Y",       False),
]


# ==================== 缓存管理 ====================
def load_cache() -> Dict:
    """加载本地缓存，返回 {metric_name: {(Year, Week_Ending, State): Value}}"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "rb") as f:
                cache = pickle.load(f)
            total = sum(len(v) for v in cache.values())
            print(f"已加载缓存: {CACHE_FILE} ({total} 条记录, {len(cache)} 个指标)")
            return cache
        except Exception as e:
            print(f"缓存加载失败 ({e})，将重新全量抓取")
    return {}


def save_cache(cache: Dict):
    """保存缓存到本地 pickle 文件"""
    with open(CACHE_FILE, "wb") as f:
        pickle.dump(cache, f)
    total = sum(len(v) for v in cache.values())
    print(f"缓存已保存: {CACHE_FILE} ({total} 条记录)")


def records_to_list(records_dict: Dict) -> List[Dict]:
    """将 {(Year, Week_Ending, State): Value} 转为记录列表"""
    result = []
    for (year, week_ending, state), value in records_dict.items():
        result.append({
            "Year": year,
            "Week_Ending": week_ending,
            "State": state,
            "Value": value
        })
    return result


# ==================== 基础抓取函数（增量版） ====================
def fetch_metric(
    metric_name: str,
    statisticcat_desc: str,
    unit_desc: str,
    state: Optional[str] = None,
    year_start: int = 2020
) -> List[Dict]:
    """
    抓取单个指标的数据（全量，不做增量过滤）。
    state: None 表示全美，否则为州代码。
    返回记录列表。
    """
    params = {
        "key": API_KEY,
        "source_desc": "SURVEY",
        "sector_desc": "CROPS",
        "group_desc": "FIELD CROPS",
        "commodity_desc": "SOYBEANS",
        "statisticcat_desc": statisticcat_desc,
        "unit_desc": unit_desc,
        "year__GE": year_start,
        "format": "JSON"
    }

    if state is None:
        params["agg_level_desc"] = "NATIONAL"
        label = "全美"
    else:
        params["agg_level_desc"] = "STATE"
        params["state_alpha"] = state
        label = f"{state} ({STATE_CODES.get(state, state)})"

    print(f"  抓取 [{metric_name}] {label} - {unit_desc} (year>={year_start})...", end=" ")

    for attempt in range(RETRY_TIMES):
        try:
            response = requests.get(BASE_URL, params=params, timeout=REQUEST_TIMEOUT)
            response.raise_for_status()
            data = response.json()

            records = []
            if "data" in data and data["data"]:
                for record in data["data"]:
                    if "year" in record and "week_ending" in record and "Value" in record:
                        val = record["Value"]
                        if val not in (None, "(D)", "", "NA"):
                            try:
                                value = float(val)
                            except (ValueError, TypeError):
                                continue
                            state_label = "US" if state is None else state
                            records.append({
                                "Year": int(record["year"]),
                                "Week_Ending": record["week_ending"],
                                "State": state_label,
                                "Value": value
                            })
            print(f"获取到 {len(records)} 条记录")
            return records
        except requests.exceptions.HTTPError as e:
            print(f"HTTP错误 (尝试 {attempt + 1}/{RETRY_TIMES}): {e}")
            if response.status_code == 400:
                print(f"    请求URL: {response.url}")
                print(f"    响应内容: {response.text[:200]}")
        except Exception as e:
            print(f"异常 (尝试 {attempt + 1}/{RETRY_TIMES}): {e}")
        time.sleep(2)

    print("最终失败，无数据返回")
    return []


def fetch_metric_incremental(
    metric_name: str,
    statisticcat_desc: str,
    unit_desc: str,
    cache: Dict,
    state: Optional[str] = None,
) -> int:
    """
    增量抓取：对比缓存，只请求缺失年份的数据，将新记录合并到缓存中。
    返回新增记录数。
    """
    # 初始化缓存条目
    if metric_name not in cache:
        cache[metric_name] = {}

    existing = cache[metric_name]
    state_key = "US" if state is None else state

    # 按州分别计算 max_year，避免某个州数据滞后被全局跳过
    state_max_year = max(
        (k[0] for k in existing.keys() if k[2] == state_key),
        default=0
    )

    if state_max_year == 0:
        year_start = 2020
    else:
        # 从该州缓存最大年份开始请求，确保不遗漏同年的新 week
        year_start = state_max_year

    raw_records = fetch_metric(metric_name, statisticcat_desc, unit_desc, state=state, year_start=year_start)

    # 只保留缓存中尚不存在的记录
    new_count = 0
    for rec in raw_records:
        key = (rec["Year"], rec["Week_Ending"], rec["State"])
        if key not in existing:
            existing[key] = rec["Value"]
            new_count += 1

    if new_count > 0:
        print(f"    -> [{metric_name}] 新增 {new_count} 条记录，缓存总计 {len(existing)} 条")
    else:
        print(f"    -> [{metric_name}] 无新数据，缓存总计 {len(existing)} 条")

    return new_count


# ==================== 优良率计算（从缓存读取，不发起网络请求） ====================
def compute_good_excellent(cache: Dict, good_key: str, excellent_key: str) -> List[Dict]:
    """从缓存中提取 Good 和 Excellent 记录，合并求和后返回记录列表"""
    good = cache.get(good_key, {})
    excellent = cache.get(excellent_key, {})

    combined = {}
    for key, val in good.items():
        combined[key] = combined.get(key, 0.0) + val
    for key, val in excellent.items():
        combined[key] = combined.get(key, 0.0) + val

    result = []
    for (year, week_ending, state), value in combined.items():
        result.append({
            "Year": year,
            "Week_Ending": week_ending,
            "State": state,
            "Value": value
        })
    return result


# ==================== 主程序 ====================
def main():
    print("=== USDA 大豆数据抓取（18州 + 全美）===")
    print("增量模式: 首次全量抓取，后续仅抓取新增数据")
    print(f"缓存文件: {CACHE_FILE}")
    print(f"输出文件: {OUTPUT_FILE}")
    print(f"指标总数: {len(METRIC_DEFS)} 个基础指标\n")

    # 1. 加载缓存
    cache = load_cache()
    total_new = 0

    # 2. 遍历所有基础指标，增量抓取
    states_to_fetch = [None] + STATES  # None = 全美，其余18州

    for metric_name, stat_cat, unit in METRIC_DEFS:
        print(f"\n--- [{metric_name}] stat_cat={stat_cat}, unit={unit} ---")

        for state in states_to_fetch:
            new = fetch_metric_incremental(metric_name, stat_cat, unit, cache, state=state)
            total_new += new
            time.sleep(REQUEST_DELAY)

    print(f"\n=== 抓取完成，共新增 {total_new} 条记录 ===")

    # 3. 保存缓存
    save_cache(cache)

    # 4. 从缓存构建 Excel 数据并写入
    print("\n=== 写入 Excel ===")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for sheet_name, base_key, avg_key, is_good_excel in SHEET_DEFS:
            if is_good_excel:
                # 优良率 = Good + Excellent
                base_records = compute_good_excellent(cache, "Good_Rate", "Excellent_Rate")
                avg_records = compute_good_excellent(cache, "Good_Rate_5Y", "Excellent_Rate_5Y")
            else:
                base_records = records_to_list(cache.get(base_key, {}))
                avg_records = records_to_list(cache.get(avg_key, {}))

            if not base_records:
                print(f"警告：{sheet_name} 无数据，创建空 Sheet")
                df = pd.DataFrame(columns=["Year", "Week_Ending", "State", "Value"])
            else:
                df = pd.DataFrame(base_records)
                df = df.sort_values(["Year", "Week_Ending", "State"])

                if avg_records:
                    avg_df = pd.DataFrame(avg_records)
                    avg_df = avg_df.sort_values(["Year", "Week_Ending", "State"])
                    avg_df["Year"] = avg_df["Year"].astype(str) + "_5Y"
                    df = pd.concat([df, avg_df], ignore_index=True)
                    print(f"  已追加 {len(avg_df)} 行五年均值到 {sheet_name}")
                else:
                    print(f"  警告: {sheet_name} 五年均值无数据")

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"已写入 Sheet: {sheet_name}，共 {len(df)} 行")

    print(f"\n数据抓取完成！文件: {OUTPUT_FILE}")
    print("每个 Sheet 包含 18州 + 全美 年度数据 + USDA官方五年均值。")


if __name__ == "__main__":
    main()
