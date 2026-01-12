# -*- coding: utf-8 -*-
"""
Measurement Analyzer - 設定常數模組
集中管理所有應用程式設定與常數
"""
from dataclasses import dataclass


@dataclass
class AppConfig:
    """應用程式設定"""
    VERSION: str = "v2.5.0"
    TITLE: str = f"量測數據分析工具 (Pro版) {VERSION}"
    LOG_FILENAME: str = "measurement_analyzer.log"
    THEME_CONFIG_FILE: str = "theme_config.txt"
    DEFAULT_TARGET_YIELD: float = 0.90  # 預設目標良率 90%
    
    class Columns:
        """資料欄位名稱"""
        FILE = '檔案名稱'
        TIME = '測量時間'
        NO = 'No'
        PROJECT = '測量專案'
        MEASURED = '實測值'
        DESIGN = '設計值'
        DIFF = '差異'
        UPPER = '上限公差'
        LOWER = '下限公差'
        RESULT = '判定結果'
        UNIT = '單位'
        ORIGINAL_JUDGE = '判斷'
        ORIGINAL_JUDGE_PDF = '判断'
        TYPE = '類型'              # [v2.5.0]
        SAMPLE_COUNT = '樣本數'    # [v2.5.0]
        NG_COUNT = 'NG數'          # [v2.5.0]
        DEFECT_RATE = '不良率(%)'  # [v2.5.0]
        CPK = 'CPK'
        SUGGESTED_TOLERANCE = '建議公差'  # [v2.3.0]
        AVERAGE = '平均值'
        MAXIMUM = '最大值'
        MINIMUM = '最小值'


# 顯示欄位列表
DISPLAY_COLUMNS = [
    AppConfig.Columns.FILE, AppConfig.Columns.TIME, AppConfig.Columns.NO, 
    AppConfig.Columns.PROJECT, AppConfig.Columns.MEASURED, AppConfig.Columns.DESIGN, 
    AppConfig.Columns.DIFF, AppConfig.Columns.UPPER, AppConfig.Columns.LOWER, 
    AppConfig.Columns.RESULT
]

# 版本更新紀錄
UPDATE_LOG = """
=== 版本更新紀錄 ===
[v2.5.0] - 2026/01/12
1. [新增] 2D 測量進階功能：
   - 2D 散佈圖 (XY Scatter Plot)、直方圖、趨勢圖
   - 2D 徑向公差 CPK (CPU) 與建議公差計算 (Rayleigh模型)
2. [新增] 陣列熱力圖 (AA區平面度)：
   - 支援顯示陣列數據的 2D 熱力圖 (所有樣本平均值)
   - 自動識別陣列測項與維度
3. [優化] 統計列表：
   - 合併顯示 2D 測項功能 (預設開啟)
   - 主表整合顯示 2D 建議公差與 CPK

[v2.4.0] - 2026/01/09
1. [重構] 程式碼模組化拆分
   - config.py: 設定常數
   - parsers.py: 檔案解析器
   - statistics.py: 統計計算
   - widgets.py: UI 元件
2. [優化] 錯誤處理強化

[v2.3.0] - 2026/01/09
1. [新增] 公差反推功能：
   - 統計表新增「建議公差(90%)」欄位
   - 雙擊測項可查看詳細公差建議分頁
   - 支援自訂目標良率 (80%~99.73%)
   - 智慧比較當前規格與建議公差
2. [依賴] 新增 scipy 套件用於統計計算

[v2.2.1] - 2025/12/04
1. [新增] 趨勢圖支援滑鼠懸停 (Hover) 功能

[v2.2.0] - 2025/12/04
1. [重構] 代碼結構優化 (Phase 3)：
   - 引入 AppConfig 集中管理全域常數，提升維護性。
   - 移除散落的常數定義 (COL_*, APP_*)。
2. [驗證] 效能基準測試：
   - 確認載入效能優異 (0.7s/100檔)，無需額外優化。

[v2.1.0] - 2025/12/04
1. [新增] CPK 可靠性分析：
   - 自動偵測小樣本 (<30)，顯示 ⚠ 警告並提供詳細 Tooltip。
   - 異常數據 (std=0) 顯示 "---"。
2. [新增] 進階自然排序 (Natural Sort)：
   - 引入 natsort 庫，完美支援 "1, 2, 10, A1, A2" 等混合編號排序。
3. [優化] 使用者介面：
   - 新增快捷鍵：Ctrl+O (開啟), Ctrl+D (清空), Ctrl+S (匯出)。
   - 進度條顯示檔案大小。
   - 表格支援像素級平滑滾動。
   - 關閉程式時的安全確認機制。
4. [修正] 系統穩定性：
   - 修正編碼問題 (UTF-8 BOM)。
   - 增強 PDF 讀取錯誤處理與日誌記錄。

[v2.0.3] - 2025/12/04
1. [新增] 自然排序功能：
   - 修正 No 欄位排序邏輯，支援 alphanumeric 排序 (如 1, 2, 10, A1, A2)。
   - 解決舊版排序錯亂問題。

[v2.0.2] - 2025/12/03 (Font Fix)
1. [修復] 圖表中文方塊問題：
   - 修正 Matplotlib Style 切換時會重置字型的問題。
   - 回歸 v1.7.1 的字型搜尋策略 (sans-serif 列表)，提高 Windows 相容性。
2. [修復] 2.0.1 崩潰問題：確認 DistributionPlotDialog 屬性正常。

[v2.0.0] - 2025/12/03
1. [重構] PDF 解析引擎：改用座標聚類法 (Coordinate Clustering) 提升精確度。
"""
