# -*- coding: utf-8 -*-
"""
Measurement Analyzer - 統計計算模組
包含 CPK 計算、建議公差反推等統計功能
"""
import numpy as np

# 嘗試匯入 scipy，若無則使用有限的 fallback
try:
    from scipy import stats as scipy_stats
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

def calculate_cpk(values, usl, lsl, min_samples=30):
    """
    計算 CPK, 添加樣本數檢查
    Returns:
        (cpk, reliability): CPK 值與可靠性標記
        reliability: 'reliable' | 'small_sample' | 'invalid'
    """
    if len(values) < 2:
        return np.nan, 'invalid'
    
    if abs(usl - lsl) < 1e-9:
        return np.nan, 'invalid'
    
    mean_val = values.mean()
    std = values.std(ddof=1)  # 使用樣本標準差
    
    if std < 1e-9:
        return 999.0, 'invalid'  # 標記為不可靠 (std=0)
    
    cpu = (usl - mean_val) / (3 * std)
    cpl = (mean_val - lsl) / (3 * std)
    cpk = min(cpu, cpl)
    
    reliability = 'reliable'
    if len(values) < min_samples:
        reliability = 'small_sample'
        
    return cpk, reliability


def calculate_tolerance_for_yield(values, design_val, target_yield=0.90):
    """
    根據目標良率反推所需公差
    
    Args:
        values: 量測值 Series
        design_val: 設計值
        target_yield: 目標良率 (0.0 ~ 1.0)
    
    Returns:
        dict: {
            'symmetric_tol': 對稱公差值,
            'upper_tol': 上限公差,
            'lower_tol': 下限公差,
            'reliability': 可靠性標記,
            'mean': 平均值,
            'std': 標準差,
            'offset': 偏移量
        }
    """
    result = {
        'symmetric_tol': np.nan,
        'upper_tol': np.nan,
        'lower_tol': np.nan,
        'reliability': 'invalid',
        'mean': np.nan,
        'std': np.nan,
        'offset': np.nan
    }
    
    if len(values) < 2:
        return result
    
    mean_val = values.mean()
    std_val = values.std(ddof=1)  # 使用樣本標準差
    
    if std_val < 1e-9:
        result['reliability'] = 'zero_std'
        result['mean'] = mean_val
        result['std'] = 0
        return result
    
    # 計算 Z 值（雙邊）
    tail_prob = (1 - target_yield) / 2
    
    if HAS_SCIPY:
        z_score = scipy_stats.norm.ppf(1 - tail_prob)
    else:
        # Fallback: 使用常見 Z 值近似
        z_table = {0.90: 1.645, 0.95: 1.96, 0.99: 2.576, 0.9973: 3.0}
        z_score = z_table.get(target_yield, 1.645)
    
    # 計算偏移量（平均值與設計值的差距）
    offset = mean_val - design_val
    
    # 對稱公差 = Z × σ + |偏移量|
    symmetric_tol = z_score * std_val + abs(offset)
    
    # 非對稱公差計算
    # 上限 = Z × σ + 偏移量（若平均偏高，上限需更大）
    # 下限 = -(Z × σ - 偏移量)
    upper_tol = z_score * std_val + offset
    lower_tol = -(z_score * std_val - offset)
    
    result['symmetric_tol'] = symmetric_tol
    result['upper_tol'] = upper_tol
    result['lower_tol'] = lower_tol
    result['mean'] = mean_val
    result['std'] = std_val
    result['offset'] = offset
    result['reliability'] = 'reliable' if len(values) >= 30 else 'small_sample'
    
    return result
