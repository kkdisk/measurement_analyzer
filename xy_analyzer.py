# -*- coding: utf-8 -*-
"""
Measurement Analyzer - XY 座標與陣列資料分析模組
包含 2D 座標配對、徑向計算、陣列資料統計

[v2.5.0] 2026/01/12
"""
import re
import numpy as np
import pandas as pd
from dataclasses import dataclass, field
from typing import Literal, Dict, List, Tuple, Optional
from enum import Enum


class MeasurementType(Enum):
    """測量資料類型"""
    SINGLE = "1D"      # 單值測項
    XY_COORD = "2D"    # XY 座標對
    ARRAY = "陣列"     # 陣列多點量測


@dataclass
class MeasurementItem:
    """單筆測量項目"""
    no: str                    # 測項編號
    project: str               # 測量專案名稱
    measured: float            # 實測值
    design: float              # 設計值
    upper_tol: float           # 上限公差
    lower_tol: float           # 下限公差
    diff: float = field(default=0.0)  # 差異值
    
    def __post_init__(self):
        self.diff = self.measured - self.design


@dataclass
class MeasurementGroup:
    """測量資料群組"""
    group_id: str                           # 群組識別碼
    group_type: MeasurementType             # 類型 (1D/2D/陣列)
    items: List[MeasurementItem] = field(default_factory=list)
    
    # 2D 專用欄位
    x_item: Optional[MeasurementItem] = None
    y_item: Optional[MeasurementItem] = None
    radial_deviation: float = np.nan        # 徑向偏差
    radial_tolerance: float = np.nan        # 徑向公差
    
    # 陣列專用欄位
    array_summary_item: Optional[MeasurementItem] = None  # [平均] 等彙總項


# ============================================================
# 解析函式
# ============================================================

def parse_xy_group(project_name: str) -> Tuple[str, str] | None:
    """
    解析測量專案名稱，提取 XY 群組 ID 與軸向
    
    支援格式：
    - 格式 A: 'NO.1_XY座標[X座標]' → ('NO.1_XY座標', 'X')
    - 格式 B: 'no.1座標[X座標]'    → ('no.1座標', 'X')
    - 格式 C: 'XY座標006[X座標]'   → ('XY座標006', 'X')
    
    Returns:
        (group_id, axis) or None
    """
    pattern = r'^(.+?)\[(X座標|Y座標)\]$'
    match = re.match(pattern, project_name)
    if match:
        group_id = match.group(1)
        axis = 'X' if 'X' in match.group(2) else 'Y'
        return (group_id, axis)
    return None


def parse_array_group(project_name: str) -> Tuple[str, int | str] | None:
    """
    解析陣列式或特殊標記的測量專案
    
    支援格式：
    - 數字索引: 'AA區平面度[5]'  → ('AA區平面度', 5)
    - 平均值:   'Gapping高度[平均]' → ('Gapping高度', '平均')
    
    Returns:
        (group_name, index_or_tag) or None
    """
    # 數字索引
    pattern_num = r'^(.+?)\[(\d+)\]$'
    match = re.match(pattern_num, project_name)
    if match:
        return (match.group(1), int(match.group(2)))
    
    # 特殊標記 (平均等)
    pattern_tag = r'^(.+?)\[(平均|最大|最小|Max|Min|Avg)\]$'
    match = re.match(pattern_tag, project_name, re.IGNORECASE)
    if match:
        return (match.group(1), match.group(2))
    
    return None


# ============================================================
# 分類函式
# ============================================================

def classify_project_name(project_name: str) -> Tuple[MeasurementType, str, Optional[str]]:
    """
    分類測量專案名稱
    
    Returns:
        (type, group_id, sub_info)
        - 1D: (SINGLE, project_name, None)
        - 2D: (XY_COORD, group_id, 'X' or 'Y')
        - 陣列: (ARRAY, group_name, index_str)
    """
    # 檢查 XY 座標
    xy_result = parse_xy_group(project_name)
    if xy_result:
        return (MeasurementType.XY_COORD, xy_result[0], xy_result[1])
    
    # 檢查陣列
    array_result = parse_array_group(project_name)
    if array_result:
        return (MeasurementType.ARRAY, array_result[0], str(array_result[1]))
    
    # 預設為單值
    return (MeasurementType.SINGLE, project_name, None)


# ============================================================
# 徑向計算函式
# ============================================================

def calculate_radial_deviation(dx: float, dy: float) -> float:
    """計算徑向偏差 (歐氏距離)"""
    return np.sqrt(dx**2 + dy**2)


def calculate_radial_tolerance(tol_x: float, tol_y: float) -> float:
    """
    計算等效徑向公差
    
    [v2.5.0] 使用 √2 倍公差（正方形對角線法）
    
    規則：
    - 若 X/Y 公差定義的是正方形區域，對角線距離 = min(tol) × √2
    - 這更符合原始規格意圖
    
    例：X/Y 各 ±0.050 → 徑向公差 = 0.050 × √2 ≈ 0.0707
    """
    if abs(tol_x) > 0 and abs(tol_y) > 0:
        # 取較小值乘以 √2
        return min(abs(tol_x), abs(tol_y)) * np.sqrt(2)
    elif abs(tol_x) > 0:
        return abs(tol_x) * np.sqrt(2)
    elif abs(tol_y) > 0:
        return abs(tol_y) * np.sqrt(2)
    return np.nan


def judge_2d_position(radial_deviation: float, radial_tolerance: float) -> str:
    """判定 2D 位置是否合格"""
    if np.isnan(radial_deviation) or np.isnan(radial_tolerance):
        return "---"
    return "OK" if radial_deviation <= radial_tolerance else "FAIL"


# ============================================================
# 資料處理函式
# ============================================================

def pair_xy_data(df: pd.DataFrame, no_col: str, project_col: str) -> Dict[str, MeasurementGroup]:
    """
    配對 XY 座標資料
    
    Args:
        df: 包含測量資料的 DataFrame
        no_col: No 欄位名稱
        project_col: 測量專案欄位名稱
    
    Returns:
        Dict[group_id, MeasurementGroup]
    """
    groups: Dict[str, MeasurementGroup] = {}
    
    for idx, row in df.iterrows():
        project = str(row.get(project_col, ''))
        type_info, group_id, sub_info = classify_project_name(project)
        
        if type_info == MeasurementType.XY_COORD:
            if group_id not in groups:
                groups[group_id] = MeasurementGroup(
                    group_id=group_id,
                    group_type=MeasurementType.XY_COORD
                )
            
            item = MeasurementItem(
                no=str(row.get(no_col, '')),
                project=project,
                measured=float(row.get('實測值', 0) or 0),
                design=float(row.get('設計值', 0) or 0),
                upper_tol=float(row.get('上限公差', 0) or 0),
                lower_tol=float(row.get('下限公差', 0) or 0)
            )
            
            if sub_info == 'X':
                groups[group_id].x_item = item
            else:
                groups[group_id].y_item = item
            
            groups[group_id].items.append(item)
    
    # 計算徑向值
    for group in groups.values():
        if group.x_item and group.y_item:
            dx = group.x_item.diff
            dy = group.y_item.diff
            group.radial_deviation = calculate_radial_deviation(dx, dy)
            group.radial_tolerance = calculate_radial_tolerance(
                group.x_item.upper_tol,
                group.y_item.upper_tol
            )
    
    return groups


def classify_all_measurements(df: pd.DataFrame, no_col: str = 'No', 
                              project_col: str = '測量專案') -> Dict[str, MeasurementGroup]:
    """
    分類所有測量資料為 1D/2D/陣列群組
    
    Returns:
        Dict[group_id, MeasurementGroup]
    """
    groups: Dict[str, MeasurementGroup] = {}
    
    # 欄位名稱 (相容多種格式)
    measured_col = '實測值'
    design_col = '設計值'
    upper_col = '上限公差'
    lower_col = '下限公差'
    
    def safe_float(val, default=0.0):
        """安全轉換為浮點數"""
        try:
            if pd.isna(val) or val == '' or val is None:
                return default
            return float(val)
        except (ValueError, TypeError):
            return default
    
    for idx, row in df.iterrows():
        project = str(row.get(project_col, ''))
        if not project or project == 'nan':
            continue
            
        type_info, group_id, sub_info = classify_project_name(project)
        
        item = MeasurementItem(
            no=str(row.get(no_col, '')),
            project=project,
            measured=safe_float(row.get(measured_col)),
            design=safe_float(row.get(design_col)),
            upper_tol=safe_float(row.get(upper_col)),
            lower_tol=safe_float(row.get(lower_col))
        )
        
        if group_id not in groups:
            groups[group_id] = MeasurementGroup(
                group_id=group_id,
                group_type=type_info
            )
        
        groups[group_id].items.append(item)
        
        # 特殊處理
        if type_info == MeasurementType.XY_COORD:
            if sub_info == 'X':
                groups[group_id].x_item = item
            else:
                groups[group_id].y_item = item
        elif type_info == MeasurementType.ARRAY and isinstance(sub_info, str):
            if sub_info in ('平均', 'Avg', 'avg'):
                groups[group_id].array_summary_item = item
    
    # 計算 2D 徑向值
    for group in groups.values():
        if group.group_type == MeasurementType.XY_COORD:
            if group.x_item and group.y_item:
                dx = group.x_item.diff
                dy = group.y_item.diff
                group.radial_deviation = calculate_radial_deviation(dx, dy)
                group.radial_tolerance = calculate_radial_tolerance(
                    group.x_item.upper_tol,
                    group.y_item.upper_tol
                )
    
    return groups


# ============================================================
# 統計函式
# ============================================================

def calculate_2d_stats(groups: List[MeasurementGroup]) -> Dict:
    """
    計算多筆 2D 資料的統計值
    
    Args:
        groups: 同一測項多次量測的 MeasurementGroup 列表
    
    Returns:
        統計結果字典
    """
    radial_devs = [g.radial_deviation for g in groups if not np.isnan(g.radial_deviation)]
    
    if not radial_devs:
        return {
            'count': 0,
            'mean_radial': np.nan,
            'max_radial': np.nan,
            'min_radial': np.nan,
            'std_radial': np.nan
        }
    
    return {
        'count': len(radial_devs),
        'mean_radial': np.mean(radial_devs),
        'max_radial': np.max(radial_devs),
        'min_radial': np.min(radial_devs),
        'std_radial': np.std(radial_devs, ddof=1) if len(radial_devs) > 1 else 0
    }


def calculate_array_stats(items: List[MeasurementItem]) -> Dict:
    """
    計算陣列資料統計（如平面度）
    
    Returns:
        統計結果字典，包含峰谷值、平均偏差等
    """
    if not items:
        return {'peak_valley': np.nan, 'mean_abs_dev': np.nan}
    
    values = [item.measured for item in items]
    
    return {
        'count': len(values),
        'peak_valley': max(values) - min(values),  # P-V 值
        'mean_abs_dev': np.mean([abs(v) for v in values]),
        'max_val': max(values),
        'min_val': min(values),
        'mean_val': np.mean(values),
        'std_val': np.std(values, ddof=1) if len(values) > 1 else 0
    }


# ============================================================
# [v2.5.0] XY 座標合併統計函式
# ============================================================

def merge_xy_stats_from_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    合併 DataFrame 中同一 XY 座標組的 X/Y 資料，計算徑向統計
    
    Args:
        df: 包含原始測量資料的 DataFrame (多檔案合併)
    
    Returns:
        合併後的統計 DataFrame，每個 XY 座標組一行
    """
    from config import AppConfig
    
    no_col = AppConfig.Columns.NO
    project_col = AppConfig.Columns.PROJECT
    measured_col = AppConfig.Columns.MEASURED
    design_col = AppConfig.Columns.DESIGN
    upper_col = AppConfig.Columns.UPPER
    lower_col = AppConfig.Columns.LOWER
    result_col = AppConfig.Columns.RESULT
    file_col = AppConfig.Columns.FILE
    
    # 收集所有 XY 座標組資料
    xy_data: Dict[str, Dict] = {}  # group_id -> {x_rows: [], y_rows: []}
    
    for idx, row in df.iterrows():
        project = str(row.get(project_col, ''))
        type_info, group_id, axis = classify_project_name(project)
        
        if type_info != MeasurementType.XY_COORD:
            continue
        
        if group_id not in xy_data:
            xy_data[group_id] = {
                'x_rows': [],
                'y_rows': [],
                'no': row.get(no_col, ''),
                'upper_tol': 0,
                'lower_tol': 0
            }
        
        # 取得公差值
        upper = pd.to_numeric(row.get(upper_col), errors='coerce')
        lower = pd.to_numeric(row.get(lower_col), errors='coerce')
        if not np.isnan(upper):
            xy_data[group_id]['upper_tol'] = upper
        if not np.isnan(lower):
            xy_data[group_id]['lower_tol'] = lower
        
        # 分類 X/Y
        if axis == 'X':
            xy_data[group_id]['x_rows'].append(row)
        else:
            xy_data[group_id]['y_rows'].append(row)
    
    # 計算每組的徑向統計
    merged_stats = []
    
    for group_id, data in xy_data.items():
        x_rows = data['x_rows']
        y_rows = data['y_rows']
        
        if not x_rows or not y_rows:
            continue
        
        # 按檔案名稱配對 X/Y 資料
        x_by_file = {row.get(file_col): row for row in x_rows}
        y_by_file = {row.get(file_col): row for row in y_rows}
        
        radial_devs = []
        ng_count = 0
        
        for file_name in x_by_file.keys():
            if file_name not in y_by_file:
                continue
            
            x_row = x_by_file[file_name]
            y_row = y_by_file[file_name]
            
            x_measured = pd.to_numeric(x_row.get(measured_col), errors='coerce')
            x_design = pd.to_numeric(x_row.get(design_col), errors='coerce')
            y_measured = pd.to_numeric(y_row.get(measured_col), errors='coerce')
            y_design = pd.to_numeric(y_row.get(design_col), errors='coerce')
            
            if any(np.isnan([x_measured, x_design, y_measured, y_design])):
                continue
            
            dx = x_measured - x_design
            dy = y_measured - y_design
            radial = calculate_radial_deviation(dx, dy)
            radial_devs.append(radial)
            
            # 判定徑向是否超標
            radial_tol = calculate_radial_tolerance(
                data['upper_tol'], 
                data['upper_tol']  # 假設 X/Y 公差相同
            )
            if radial > radial_tol:
                ng_count += 1
        
        if not radial_devs:
            continue
        
        # 統計結果
        merged_stats.append({
            'No': data['no'],
            '測量專案': group_id,
            '類型': '2D',
            '樣本數': len(radial_devs),
            'NG數': ng_count,
            '徑向偏差_平均': np.mean(radial_devs),
            '徑向偏差_最大': np.max(radial_devs),
            '徑向偏差_最小': np.min(radial_devs),
            '徑向公差': calculate_radial_tolerance(data['upper_tol'], data['upper_tol']),
            '_x_project': f"{group_id}[X座標]",
            '_y_project': f"{group_id}[Y座標]"
        })
    
    return pd.DataFrame(merged_stats)


def get_xy_group_id(project_name: str) -> Optional[str]:
    """
    獲取 XY 座標組的 group_id
    
    Returns:
        group_id 或 None (若非 XY 座標)
    """
    result = parse_xy_group(project_name)
    return result[0] if result else None


def calculate_2d_suggested_tolerance(radial_vals: np.ndarray, target_yield: float = 0.90) -> Dict[str, float]:
    """
    計算 2D 建議公差 (基於 Rayleigh 分佈假設)
    
    假設 X, Y 誤差為獨立且變異數相同的常態分佈，且平均值為 0 (理想同心)
    則徑向偏差 R = sqrt(X^2 + Y^2) 服從 Rayleigh 分佈
    
    參數估計：
    sigma ≈ sqrt( mean(R^2) / 2 )  (使用均方根估計，包含偏移影響)
    
    建議公差 T 滿足 P(R <= T) = target_yield
    T = sigma * sqrt(-2 * ln(1 - target_yield))
    
    Args:
        radial_vals: 徑向偏差陣列
        target_yield: 目標良率 (0.0 ~ 1.0)
        
    Returns:
        {
            'suggested_tol': 建議公差值,
            'sigma': 估計的 Rayleigh 參數,
            'reliability': 可靠度說明
        }
    """
    if len(radial_vals) < 2:
        return {'suggested_tol': np.nan, 'sigma': np.nan, 'reliability': 'insufficient_data'}
    
    # 清除 NaN
    radial_vals = radial_vals[~np.isnan(radial_vals)]
    if len(radial_vals) == 0:
        return {'suggested_tol': np.nan, 'sigma': np.nan, 'reliability': 'no_data'}

    # 1. 參數估計 (使用 RMS 方法，這樣可以包含中心偏移的影響)
    # sigma_hat = sqrt( sum(r^2) / (2n) )
    mean_sq = np.mean(radial_vals**2)
    sigma = np.sqrt(mean_sq / 2)
    
    if sigma <= 0:
        return {'suggested_tol': 0.0, 'sigma': 0.0, 'reliability': 'zero_variance'}
        
    # 2. 計算對應良率的公差
    # Formula: T = sigma * sqrt(-2 * ln(1 - P))
    try:
        suggested_tol = sigma * np.sqrt(-2 * np.log(1 - target_yield))
    except:
        suggested_tol = np.nan
        
    reliability = 'ok'
    if len(radial_vals) < 30:
        reliability = 'small_sample'
        
    return {
        'suggested_tol': suggested_tol,
        'sigma': sigma,
        'reliability': reliability,
        'target_yield': target_yield
    }


