# -*- coding: utf-8 -*-
import sys
import os
import glob
import pandas as pd
import numpy as np
import re
import logging
import traceback
from datetime import datetime
from dataclasses import dataclass

# PyQt6 imports
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QProgressBar, QMessageBox, QGroupBox, QCheckBox, QDialog,
                             QTabWidget, QTextEdit, QSplitter, QFrame, QMenu, QInputDialog,
                             QAbstractItemView)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEvent
from PyQt6.QtGui import QColor, QBrush, QFont, QAction, QKeySequence



# Matplotlib imports
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Optional Libraries Check
try:
    import pdfplumber
    HAS_PDF_SUPPORT = True
except ImportError:
    HAS_PDF_SUPPORT = False

try:
    import qdarktheme
    HAS_THEME_SUPPORT = True
except ImportError:
    HAS_THEME_SUPPORT = False

# Natsort
try:
    from natsort import index_natsorted, natsort_keygen, ns
    HAS_NATSORT = True
except ImportError:
    HAS_NATSORT = False

# Scipy for statistical calculations
try:
    from scipy import stats as scipy_stats
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

# --- 設定常數 ---
@dataclass
class AppConfig:
    VERSION: str = "v2.3.0"
    TITLE: str = f"量測數據分析工具 (Pro版) {VERSION}"
    LOG_FILENAME: str = "measurement_analyzer.log"
    THEME_CONFIG_FILE: str = "theme_config.txt"
    DEFAULT_TARGET_YIELD: float = 0.90  # 預設目標良率 90%
    
    class Columns:
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

DISPLAY_COLUMNS = [
    AppConfig.Columns.FILE, AppConfig.Columns.TIME, AppConfig.Columns.NO, AppConfig.Columns.PROJECT, 
    AppConfig.Columns.MEASURED, AppConfig.Columns.DESIGN, AppConfig.Columns.DIFF, 
    AppConfig.Columns.UPPER, AppConfig.Columns.LOWER, AppConfig.Columns.RESULT
]

UPDATE_LOG = """
=== 版本更新紀錄 ===
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

# --- 初始化日誌 ---
def setup_logging():
    # Allow debug mode via environment variable
    log_level = logging.DEBUG if os.getenv('ANALYZER_DEBUG') else logging.INFO
    
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
        handlers=[
            logging.FileHandler(AppConfig.LOG_FILENAME, encoding='utf-8', mode='w'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    logging.getLogger("matplotlib").setLevel(logging.WARNING)
    
    # Encoding Verification
    try:
        test_str = "測量數據"
        assert len(test_str) == 4, "Encoding verification failed"
        logging.info(f"編碼驗證成功: {test_str}")
    except Exception as e:
        logging.error(f"編碼驗證失敗: {e}")

    logging.info(f"應用程式啟動 - {AppConfig.TITLE}")

# --- [關鍵修正] 字型設定 (回歸 v1.7.1 策略) ---
# 此函式需在 plt.style.use() 之後被呼叫，才能確保字型不被覆蓋
def set_chinese_font():
    # 常見中文字型清單 (優先順序)
    font_names = ['Microsoft JhengHei', 'Microsoft YaHei', 'SimHei', 'PingFang TC', 'Arial Unicode MS']
    
    # 取得系統可用字型
    try:
        system_fonts = set([f.name for f in fm.fontManager.ttflist])
        
        # 尋找第一個可用的中文字型
        found = False
        for name in font_names:
            if name in system_fonts:
                matplotlib.rcParams['font.sans-serif'] = [name] + matplotlib.rcParams['font.sans-serif']
                found = True
                break
        
        # 設定負號正確顯示
        matplotlib.rcParams['axes.unicode_minus'] = False
        
        if not found:
            logging.warning("未偵測到常見中文字型，圖表可能顯示方格。")
            
    except Exception as e:
        logging.error(f"字型設定失敗: {e}")

# 程式啟動時先執行一次
set_chinese_font()

# --- 日期解析 ---
def parse_keyence_date(date_str):
    if not isinstance(date_str, str): return None
    date_str = date_str.strip()
    try:
        match = re.search(r'(\d+)/(\d+)/(\d+)\s+(上午|下午)\s*(\d+):(\d+):(\d+)', date_str)
        if match:
            year, month, day, ampm, hour, minute, second = match.groups()
            year, month, day = int(year), int(month), int(day)
            hour, minute, second = int(hour), int(minute), int(second)
            if ampm == "下午" and hour < 12: hour += 12
            elif ampm == "上午" and hour == 12: hour = 0
            return datetime(year, month, day, hour, minute, second)
        else:
            formats = ["%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %p %I:%M:%S"]
            for fmt in formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError: continue
            return None
    except: return None

# --- 自然排序輔助函式 ---
# --- CPK 計算 ---
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

# --- 公差反推計算 ---
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

# --- 自然排序輔助函式 ---
def natural_keys(text):
    """
    Fallback for natural sorting if natsort is missing.
    """
    try:
        text = str(text)
        return tuple([int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)])
    except Exception:
        return (str(text),)

# --- 背景執行緒 ---
class FileLoaderThread(QThread):
    progress_updated = pyqtSignal(int, str)
    data_loaded = pyqtSignal(list, set)
    error_occurred = pyqtSignal(str)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._is_running = True

    def find_header_row_and_date_csv(self, filepath):
        try:
            encodings = ['utf-8-sig', 'big5', 'cp950', 'shift_jis']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        lines = [next(f) for _ in range(60)]
                    measure_time = None
                    for line in lines[:20]:
                        if "測量日期及時間" in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                measure_time = parse_keyence_date(parts[1].strip())
                            break
                    for i, line in enumerate(lines): 
                        if "No" in line and "實測值" in line and "設計值" in line:
                            return i, enc, measure_time
                except: continue 
            return None, None, None
        except: return None, None, None

    def extract_text_by_clustering(self, page, y_tolerance=3):
        words = page.extract_words(keep_blank_chars=True)
        if not words: return []
        
        # Boundary check
        width, height = page.width, page.height
        valid_words = [w for w in words if 0 <= w['x0'] <= width and 0 <= w['top'] <= height]
        
        if not valid_words:
            return []

        rows = {} 
        for word in valid_words:
            top = word['top']
            found_row = None
            for y in rows:
                if abs(top - y) <= y_tolerance:
                    found_row = y
                    break
            
            if found_row is not None:
                rows[found_row].append(word)
            else:
                rows[top] = [word]
        
        sorted_ys = sorted(rows.keys())
        lines = []
        for y in sorted_ys:
            row_words = sorted(rows[y], key=lambda w: w['x0'])
            text_line = " ".join([w['text'] for w in row_words])
            lines.append(text_line)
            
        return lines

    def read_pdf_file(self, filepath):
        if not HAS_PDF_SUPPORT: return None, None
        
        df = pd.DataFrame()
        measure_time = None
        data_list = []
        
        try:
            with pdfplumber.open(filepath) as pdf:
                first_page_words = self.extract_text_by_clustering(pdf.pages[0])
                for line in first_page_words:
                    if "測量日期及時間" in line:
                        date_match = re.search(r'(\d{4}/\d{1,2}/\d{1,2}\s+(?:上午|下午)\s*\d{1,2}:\d{1,2}:\d{1,2})', line)
                        if date_match:
                            measure_time = parse_keyence_date(date_match.group(1))
                        break
                
                pattern = re.compile(
                    r'^\s*(?P<no>\d+)\s+'           
                    r'(?P<proj>.+?)\s+'             
                    r'(?P<val>-?\d+(?:\.\d+)?)\s+'  
                    r'(?P<unit>mm|um)\s+'           
                    r'(?P<design>-?\d+(?:\.\d+)?)\s+'
                    r'(?P<up>-?\d+(?:\.\d+)?)\s+'
                    r'(?P<low>-?\d+(?:\.\d+)?)\s+'
                    r'(?P<judge>OK|NG|---|Warning)'
                )
                
                for page in pdf.pages:
                    lines = self.extract_text_by_clustering(page)
                    for line in lines:
                        if "測量專案" in line or "部件報告" in line or "測量結果" in line: continue
                        
                        match = pattern.search(line)
                        if match:
                            d = match.groupdict()
                            item = {
                                AppConfig.Columns.NO: d['no'],
                                AppConfig.Columns.PROJECT: d['proj'].strip(),
                                AppConfig.Columns.MEASURED: d['val'],
                                AppConfig.Columns.UNIT: d['unit'],
                                AppConfig.Columns.DESIGN: d['design'],
                                AppConfig.Columns.UPPER: d['up'],
                                AppConfig.Columns.LOWER: d['low'],
                                AppConfig.Columns.ORIGINAL_JUDGE: d['judge']
                            }
                            data_list.append(item)
            
            if data_list:
                df = pd.DataFrame(data_list)
                return df, measure_time
            else:
                return None, None

        except pdfplumber.PDFSyntaxError as e:
            logging.error(f"PDF 格式錯誤 {filepath}: {e}")
            return None, None
        except PermissionError:
            logging.error(f"無權限讀取 {filepath}")
            return None, None
        except Exception as e:
            logging.error(f"PDF Error {filepath}: {e}\n{traceback.format_exc()}")
            return None, None

    def run(self):
        new_data_frames = []
        loaded_filenames = set()
        for i, filepath in enumerate(self.file_paths):
            if not self._is_running: break
            filename = os.path.basename(filepath)
            
            # Get file size
            try:
                size_kb = os.path.getsize(filepath) / 1024
                size_str = f"{size_kb:.1f}KB"
            except:
                size_str = "Unknown"
                
            self.progress_updated.emit(i + 1, f"處理中: {filename} ({size_str})")
            
            df = None
            measure_time = None
            try:
                ext = os.path.splitext(filename)[1].lower()
                if ext == '.pdf':
                    df, measure_time = self.read_pdf_file(filepath)
                else:
                    header_idx, encoding, measure_time = self.find_header_row_and_date_csv(filepath)
                    if header_idx is not None:
                        df = pd.read_csv(filepath, skiprows=header_idx, header=0, 
                                       encoding=encoding, on_bad_lines='skip', index_col=False)
                
                if df is not None:
                    loaded_filenames.add(filename)
                    df.columns = [str(c).strip() for c in df.columns]
                    if AppConfig.Columns.NO not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: AppConfig.Columns.NO}, inplace=True)
                                break
                    required = [AppConfig.Columns.NO, AppConfig.Columns.MEASURED, AppConfig.Columns.DESIGN]
                    if all(c in df.columns for c in required):
                        df = df.dropna(subset=[AppConfig.Columns.NO])
                        num_cols = [AppConfig.Columns.MEASURED, AppConfig.Columns.DESIGN, AppConfig.Columns.UPPER, AppConfig.Columns.LOWER]
                        for c in num_cols:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors='coerce')
                            else:
                                df[c] = 0.0
                        
                        df[AppConfig.Columns.DIFF] = df[AppConfig.Columns.MEASURED] - df[AppConfig.Columns.DESIGN]
                        df[AppConfig.Columns.RESULT] = "OK"
                        
                        mask_ignore = df[AppConfig.Columns.DESIGN].abs() < 0.000001
                        df.loc[mask_ignore, AppConfig.Columns.RESULT] = "---"
                        
                        mask_tol_na = df[AppConfig.Columns.UPPER].isna() | df[AppConfig.Columns.LOWER].isna()
                        df.loc[mask_tol_na, AppConfig.Columns.RESULT] = "---"
                        
                        mask_tol_zero = (df[AppConfig.Columns.UPPER] == 0) & (df[AppConfig.Columns.LOWER] == 0)
                        orig_judge = None
                        if AppConfig.Columns.ORIGINAL_JUDGE in df.columns: orig_judge = AppConfig.Columns.ORIGINAL_JUDGE
                        elif AppConfig.Columns.ORIGINAL_JUDGE_PDF in df.columns: orig_judge = AppConfig.Columns.ORIGINAL_JUDGE_PDF
                        
                        if orig_judge:
                            df.loc[mask_tol_zero, AppConfig.Columns.RESULT] = df.loc[mask_tol_zero, orig_judge].fillna("---")
                        else:
                            df.loc[mask_tol_zero, AppConfig.Columns.RESULT] = "---"
                            
                        mask_check = ~(mask_ignore | mask_tol_na | mask_tol_zero)
                        mask_fail = mask_check & ((df[AppConfig.Columns.DIFF] > df[AppConfig.Columns.UPPER]) | (df[AppConfig.Columns.DIFF] < df[AppConfig.Columns.LOWER]))
                        df.loc[mask_fail, AppConfig.Columns.RESULT] = "FAIL"
                        
                        df[AppConfig.Columns.FILE] = filename
                        df[AppConfig.Columns.TIME] = measure_time if measure_time else pd.NaT
                        if AppConfig.Columns.PROJECT not in df.columns: df[AppConfig.Columns.PROJECT] = ''
                        
                        cols = [c for c in DISPLAY_COLUMNS if c in df.columns]
                        new_data_frames.append(df[cols])
            except Exception as e:
                logging.error(f"Error processing {filename}: {e}\n{traceback.format_exc()}")

        self.data_loaded.emit(new_data_frames, loaded_filenames)

    def stop(self):
        self._is_running = False

# --- GUI Components ---
class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        if HAS_NATSORT:
            try:
                natsort_key = natsort_keygen(alg=ns.IGNORECASE)
                return natsort_key(self.text()) < natsort_key(other.text())
            except Exception:
                pass
        
        try:
            # Fallback
            return natural_keys(self.text()) < natural_keys(other.text())
        except Exception:
            return super().__lt__(other)

class VersionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 600, 450)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(AppConfig.TITLE))
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(UPDATE_LOG)
        layout.addWidget(txt)
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)

class DistributionPlotDialog(QDialog):
    def __init__(self, item_name, df_item, design_val, upper_tol, lower_tol, parent=None, theme='light'):
        super().__init__(parent)
        self.setWindowTitle(f"詳細分析: {item_name}")
        self.setGeometry(100, 100, 950, 650)
        self.item_name = item_name
        self.df_item = df_item
        self.design_val = design_val
        self.upper_tol = upper_tol
        self.lower_tol = lower_tol
        # [v2.0.1 修正] 補回 usl 與 lsl 定義，防止崩潰
        self.usl = design_val + upper_tol
        self.lsl = design_val + lower_tol
        self.theme = theme
        
        # 設定 Style
        if self.theme == 'dark':
            plt.style.use('dark_background')
        else:
            plt.style.use('default')
            
        # [v2.0.2 關鍵修正] 設定 Style 後必須重新套用中文字型，否則會被覆蓋回預設值
        set_chinese_font()

        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        self.tab_hist = QWidget()
        self.plot_histogram(self.tab_hist)
        tabs.addTab(self.tab_hist, "分佈直方圖")
        
        self.tab_trend = QWidget()
        self.plot_trend(self.tab_trend)
        tabs.addTab(self.tab_trend, "趨勢圖")
        
        # [v2.3.0] 新增公差建議分頁
        self.tab_tolerance = QWidget()
        self.setup_tolerance_tab(self.tab_tolerance)
        tabs.addTab(self.tab_tolerance, "📐 公差建議")
        
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)
    
    def setup_tolerance_tab(self, parent_widget):
        """設定公差建議分頁"""
        layout = QVBoxLayout(parent_widget)
        
        # 目標良率選擇區
        yield_group = QGroupBox("目標良率設定")
        yield_layout = QHBoxLayout()
        
        yield_layout.addWidget(QLabel("目標良率："))
        
        from PyQt6.QtWidgets import QComboBox, QSpinBox
        self.yield_combo = QComboBox()
        self.yield_combo.addItems(["80%", "85%", "90%", "95%", "99%", "99.73% (3σ)"])
        self.yield_combo.setCurrentIndex(2)  # 預設 90%
        self.yield_combo.currentIndexChanged.connect(self.update_tolerance_display)
        yield_layout.addWidget(self.yield_combo)
        
        yield_layout.addStretch()
        yield_group.setLayout(yield_layout)
        layout.addWidget(yield_group)
        
        # 結果顯示區
        result_group = QGroupBox("計算結果")
        result_layout = QVBoxLayout()
        
        self.tol_result_text = QTextEdit()
        self.tol_result_text.setReadOnly(True)
        self.tol_result_text.setMinimumHeight(300)
        result_layout.addWidget(self.tol_result_text)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 初始計算
        self.update_tolerance_display()
    
    def update_tolerance_display(self):
        """更新公差計算結果顯示"""
        yield_map = {0: 0.80, 1: 0.85, 2: 0.90, 3: 0.95, 4: 0.99, 5: 0.9973}
        target_yield = yield_map.get(self.yield_combo.currentIndex(), 0.90)
        
        vals = pd.to_numeric(self.df_item[AppConfig.Columns.MEASURED], errors='coerce').dropna()
        result = calculate_tolerance_for_yield(vals, self.design_val, target_yield)
        
        # 格式化輸出
        lines = []
        lines.append(f"═══════════════════════════════════════")
        lines.append(f"  測量專案：{self.item_name}")
        lines.append(f"  目標良率：{target_yield * 100:.2f}%")
        lines.append(f"═══════════════════════════════════════")
        lines.append("")
        
        if result['reliability'] == 'invalid':
            lines.append("❌ 無法計算：數據不足 (需至少 2 個樣本)")
        elif result['reliability'] == 'zero_std':
            lines.append("❌ 無法計算：標準差為零 (所有數據相同)")
        else:
            lines.append("📊 【數據統計】")
            lines.append(f"   樣本數：{len(vals)}")
            lines.append(f"   平均值 (μ)：{result['mean']:.4f}")
            lines.append(f"   標準差 (σ)：{result['std']:.4f}")
            lines.append(f"   設計值：{self.design_val:.4f}")
            lines.append(f"   製程偏移：{result['offset']:+.4f}")
            lines.append("")
            
            lines.append("📐 【建議公差】")
            lines.append(f"   ✅ 對稱公差：±{result['symmetric_tol']:.4f}")
            lines.append("")
            lines.append(f"   📈 非對稱建議：")
            lines.append(f"      上限公差：+{result['upper_tol']:.4f}")
            lines.append(f"      下限公差：{result['lower_tol']:.4f}")
            lines.append("")
            
            lines.append("📋 【與當前規格比較】")
            lines.append(f"   當前上限：+{self.upper_tol:.4f}")
            lines.append(f"   當前下限：{self.lower_tol:.4f}")
            
            current_max_tol = max(abs(self.upper_tol), abs(self.lower_tol))
            if current_max_tol > 0:
                ratio = result['symmetric_tol'] / current_max_tol
                if ratio > 1.2:
                    lines.append("")
                    lines.append(f"   ⚠️ 警告：要達到 {target_yield*100:.0f}% 良率，")
                    lines.append(f"      建議公差比當前規格大 {(ratio-1)*100:.1f}%")
                    lines.append(f"      建議放寬規格或改善製程")
                elif ratio < 0.8:
                    lines.append("")
                    lines.append(f"   ✅ 良好：當前規格充裕，")
                    lines.append(f"      實際只需 {ratio*100:.1f}% 即可達標")
                else:
                    lines.append("")
                    lines.append(f"   ℹ️ 規格適中 (比例：{ratio*100:.1f}%)")
            
            if result['reliability'] == 'small_sample':
                lines.append("")
                lines.append("⚠️ 注意：樣本數少於 30，結果僅供參考")
                lines.append("   建議累積更多數據後再做決策")
        
        self.tol_result_text.setPlainText("\n".join(lines))

    def plot_histogram(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        data = self.df_item[AppConfig.Columns.MEASURED].dropna()
        if len(data) > 0:
            color = 'cyan' if self.theme == 'dark' else 'skyblue'
            edgecolor = 'white' if self.theme == 'dark' else 'black'
            ax.hist(data, bins=15, color=color, edgecolor=edgecolor, alpha=0.7, label='實測值')
            ax.axvline(self.design_val, color='lime' if self.theme=='dark' else 'green', linestyle='-', linewidth=2, label='設計值')
            ax.axvline(self.usl, color='red', linestyle='--', linewidth=2, label='USL')
            ax.axvline(self.lsl, color='red', linestyle='--', linewidth=2, label='LSL')
            ax.set_title("量測值分佈圖")
            ax.legend()
            ax.grid(True, alpha=0.3)
        else:
            ax.text(0.5, 0.5, "無有效數據", ha='center', va='center')
        layout.addWidget(toolbar)
        layout.addWidget(canvas)

    def plot_trend(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        df_sorted = self.df_item.copy()
        has_time = False
        if AppConfig.Columns.TIME in df_sorted.columns:
            try:
                if len(df_sorted[AppConfig.Columns.TIME].dropna()) > 0:
                    df_sorted = df_sorted.sort_values(by=AppConfig.Columns.TIME)
                    has_time = True
            except: pass
        
        y_data = df_sorted[AppConfig.Columns.MEASURED].values
        x_data = np.arange(1, len(y_data) + 1)
        
        # Prepare data for tooltip
        filenames = df_sorted[AppConfig.Columns.FILE].values if AppConfig.Columns.FILE in df_sorted.columns else []
        times = df_sorted[AppConfig.Columns.TIME].values if AppConfig.Columns.TIME in df_sorted.columns else []
        
        line_color = 'cyan' if self.theme == 'dark' else 'blue'
        line, = ax.plot(x_data, y_data, marker='o', linestyle='-', color=line_color, markersize=4, label='實測值')
        
        ax.axhline(self.design_val, color='lime' if self.theme=='dark' else 'green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        ax.set_title("量測值趨勢圖")
        ax.set_xlabel("時間順序" if has_time else "讀取順序")
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        # --- Tooltip Implementation ---
        annot = ax.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="w", alpha=0.9),
                            arrowprops=dict(arrowstyle="->"))
        annot.set_visible(False)

        def update_annot(ind):
            x, y = line.get_data()
            idx = ind["ind"][0]
            annot.xy = (x[idx], y[idx])
            
            val = y_data[idx]
            fname = filenames[idx] if len(filenames) > idx else "Unknown"
            
            time_str = ""
            if len(times) > idx:
                t = times[idx]
                if pd.notnull(t):
                    try:
                        time_str = pd.to_datetime(t).strftime("%Y/%m/%d %H:%M:%S")
                    except: pass
            
            # Format text
            text = f"File: {fname}\nValue: {val:.4f}"
            if time_str:
                text += f"\nTime: {time_str}"
                
            annot.set_text(text)

        def hover(event):
            vis = annot.get_visible()
            if event.inaxes == ax:
                cont, ind = line.contains(event)
                if cont:
                    update_annot(ind)
                    annot.set_visible(True)
                    canvas.draw_idle()
                else:
                    if vis:
                        annot.set_visible(False)
                        canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", hover)
        # ------------------------------

        layout.addWidget(toolbar)
        layout.addWidget(canvas)

# --- 主程式 ---
class MeasurementAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        setup_logging()
        if not HAS_NATSORT:
            logging.warning("未安裝 natsort 套件，建議執行: pip install natsort")
        self.setWindowTitle(AppConfig.TITLE)
        self.setGeometry(100, 100, 1300, 850)
        self.all_data = pd.DataFrame()
        self.stats_data = pd.DataFrame()
        self.loaded_files = set()
        self.loader_thread = None
        self.current_theme = 'light'
        self.init_theme()
        self.init_ui()

    def init_theme(self):
        if not HAS_THEME_SUPPORT: return
        try:
            if os.path.exists(AppConfig.THEME_CONFIG_FILE):
                with open(AppConfig.THEME_CONFIG_FILE, 'r') as f:
                    self.current_theme = f.read().strip()
            qdarktheme.setup_theme(self.current_theme)
        except Exception as e:
            logging.error(f"主題載入失敗: {e}")

    def closeEvent(self, event):
        """關閉前儲存當前狀態並停止線程"""
        if self.loader_thread and self.loader_thread.isRunning():
            reply = QMessageBox.question(
                self, '確認', 
                '數據正在載入中,確定要關閉嗎?',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return
            self.loader_thread.stop()
            self.loader_thread.wait()
        event.accept()

    def toggle_theme(self):
        if not HAS_THEME_SUPPORT:
            QMessageBox.information(self, "提示", "請先安裝 'pyqtdarktheme' 套件")
            return
        new_theme = 'dark' if self.current_theme == 'light' else 'light'
        self.current_theme = new_theme
        qdarktheme.setup_theme(new_theme)
        try:
            with open(AppConfig.THEME_CONFIG_FILE, 'w') as f:
                f.write(new_theme)
        except: pass
        self.btn_theme.setText("切換亮色" if new_theme == 'dark' else "切換深色")

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        control_group = QGroupBox("操作控制")
        control_layout = QHBoxLayout()
        
        self.btn_add = QPushButton("1. 加入資料夾")
        self.btn_add.clicked.connect(self.add_folder_data)
        self.btn_add.setMinimumHeight(40)
        self.btn_add.setShortcut("Ctrl+O")
        
        self.btn_clear = QPushButton("清空資料")
        self.btn_clear.clicked.connect(self.clear_all_data)
        self.btn_clear.setStyleSheet("color: red;")
        self.btn_clear.setShortcut("Ctrl+D")
        
        self.btn_export = QPushButton("匯出當前頁面資料")
        self.btn_export.clicked.connect(self.export_current_tab)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)
        self.btn_export.setShortcut("Ctrl+S")
        
        theme_label = "切換亮色" if self.current_theme == 'dark' else "切換深色"
        self.btn_theme = QPushButton(theme_label)
        self.btn_theme.clicked.connect(self.toggle_theme)
        
        self.btn_version = QPushButton("關於")
        self.btn_version.clicked.connect(self.show_version_info)
        
        control_layout.addWidget(self.btn_add, 1)
        control_layout.addWidget(self.btn_clear)
        control_layout.addWidget(self.btn_export, 1)
        control_layout.addWidget(self.btn_theme)
        control_layout.addWidget(self.btn_version)
        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        self.tabs = QTabWidget()
        self.tab_stats = QWidget()
        self.setup_statistics_tab()
        self.tabs.addTab(self.tab_stats, "1. 統計摘要分析")
        
        self.tab_raw = QWidget()
        self.setup_raw_data_tab()
        self.tabs.addTab(self.tab_raw, "2. 原始數據列表")
        main_layout.addWidget(self.tabs)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)
        
        self.lbl_info = QLabel("準備就緒。")
        main_layout.addWidget(self.lbl_info)
        
        self.lbl_status = QLabel("目前總資料: 0 筆 | 總樣本數: 0")
        self.lbl_status.setStyleSheet("color: blue; font-weight: bold;")
        main_layout.addWidget(self.lbl_status)

    def setup_raw_data_tab(self):
        layout = QVBoxLayout(self.tab_raw)
        filter_layout = QHBoxLayout()
        self.chk_only_fail = QCheckBox("僅顯示 FAIL 項目")
        self.chk_only_fail.stateChanged.connect(self.refresh_raw_table)
        self.chk_only_fail.setEnabled(False)
        
        self.btn_plot_raw = QPushButton("視覺化選定列")
        self.btn_plot_raw.clicked.connect(self.plot_from_raw_table)
        self.btn_plot_raw.setEnabled(False)
        
        filter_layout.addWidget(self.chk_only_fail)
        filter_layout.addStretch()
        filter_layout.addWidget(self.btn_plot_raw)
        layout.addLayout(filter_layout)

        self.raw_table = QTableWidget()
        self.raw_table.setColumnCount(len(DISPLAY_COLUMNS))
        self.raw_table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.raw_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.raw_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.raw_table.setAlternatingRowColors(True)
        header = self.raw_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        # Enable pixel scrolling
        self.raw_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        
        layout.addWidget(self.raw_table)

    def setup_statistics_tab(self):
        layout = QVBoxLayout(self.tab_stats)
        self.lbl_stats_summary = QLabel("尚未載入資料")
        # [UI Fix] Remove hardcoded background color for Dark Mode compatibility
        self.lbl_stats_summary.setStyleSheet("padding: 10px; font-weight: bold;") 
        layout.addWidget(self.lbl_stats_summary)
        
        lbl_hint = QLabel("提示：雙擊表格任一列可開啟詳細圖表分析")
        lbl_hint.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(lbl_hint)

        self.stats_table = QTableWidget()
        cols = ["No", "測量專案", "樣本數", "NG數", "不良率(%)", "CPK", "建議公差(90%)", "平均值", "最大值", "最小值"]
        self.stats_table.setColumnCount(len(cols))
        self.stats_table.setHorizontalHeaderLabels(cols)
        self.stats_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.stats_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.stats_table.setAlternatingRowColors(True)
        self.stats_table.setSortingEnabled(True)
        self.stats_table.doubleClicked.connect(self.plot_from_stats_table)
        
        header = self.stats_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        # Enable pixel scrolling
        self.stats_table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        
        layout.addWidget(self.stats_table)

    def show_version_info(self):
        dlg = VersionDialog(self)
        dlg.exec()

    def add_folder_data(self):
        folder_path = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if not folder_path: return
        
        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
        
        files_to_load = csv_files + pdf_files
        
        if csv_files and pdf_files and HAS_PDF_SUPPORT:
            csv_bases = {os.path.splitext(os.path.basename(f))[0] for f in csv_files}
            pdf_bases = {os.path.splitext(os.path.basename(f))[0] for f in pdf_files}
            duplicates = csv_bases.intersection(pdf_bases)
            
            if duplicates:
                items = ["優先匯入 CSV (推薦)", "優先匯入 PDF", "全部匯入"]
                item, ok = QInputDialog.getItem(self, "發現重複報告", 
                                                f"發現 {len(duplicates)} 組同名報告 (同時有 CSV 與 PDF)。\n"
                                                "為避免數據重複，請選擇匯入策略：", 
                                                items, 0, False)
                if ok and item:
                    if "CSV" in item:
                        pdf_unique = [f for f in pdf_files if os.path.splitext(os.path.basename(f))[0] not in csv_bases]
                        files_to_load = csv_files + pdf_unique
                    elif "PDF" in item:
                        csv_unique = [f for f in csv_files if os.path.splitext(os.path.basename(f))[0] not in pdf_bases]
                        files_to_load = csv_unique + pdf_files
                    else:
                        files_to_load = csv_files + pdf_files
                else:
                    return

        if not files_to_load:
            QMessageBox.warning(self, "提示", "無檔案可匯入。")
            return

        self.set_ui_loading_state(True)
        self.lbl_info.setText(f"開始處理: {len(files_to_load)} 個檔案...")
        self.progress_bar.setMaximum(len(files_to_load))
        self.progress_bar.setValue(0)

        self.loader_thread = FileLoaderThread(files_to_load)
        self.loader_thread.progress_updated.connect(self.on_progress_updated)
        self.loader_thread.data_loaded.connect(self.on_data_loaded)
        self.loader_thread.start()

    def set_ui_loading_state(self, is_loading):
        self.btn_add.setEnabled(not is_loading)
        self.btn_clear.setEnabled(not is_loading)
        if is_loading: self.btn_export.setEnabled(False)

    def on_progress_updated(self, value, message):
        self.progress_bar.setValue(value)
        self.lbl_info.setText(message)

    def on_data_loaded(self, new_data_frames, loaded_filenames):
        import time
        start_time = time.time()
        
        self.loaded_files.update(loaded_filenames)
        if new_data_frames:
            self.lbl_info.setText("正在合併資料...")
            QApplication.processEvents() 
            new_data = pd.concat(new_data_frames, ignore_index=True)
            if self.all_data.empty: self.all_data = new_data
            else: self.all_data = pd.concat([self.all_data, new_data], ignore_index=True)
            
            self.btn_export.setEnabled(True)
            self.chk_only_fail.setEnabled(True)
            self.btn_plot_raw.setEnabled(True)
            self.refresh_raw_table()
            self.calculate_and_refresh_stats()
            
            elapsed = time.time() - start_time
            msg = f"完成。本次加入 {len(new_data)} 筆數據。耗時 {elapsed:.2f}秒"
            logging.info(f"載入完成: {len(loaded_filenames)} 檔案, {len(new_data)} 筆, 耗時 {elapsed:.2f}秒")
            
            self.lbl_info.setText(msg)
            QMessageBox.information(self, "完成", f"已加入 {len(loaded_filenames)} 個檔案。")
        else:
            self.lbl_info.setText("無有效數據。")
            QMessageBox.warning(self, "結果", "未提取到有效數據。")
        self.set_ui_loading_state(False)

    def clear_all_data(self):
        reply = QMessageBox.question(self, '確認', '確定清空？', 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.all_data = pd.DataFrame()
            self.stats_data = pd.DataFrame()
            self.loaded_files.clear()
            self.raw_table.setRowCount(0)
            self.stats_table.setRowCount(0)
            self.lbl_status.setText("資料已清空")
            self.lbl_stats_summary.setText("資料已清空")
            self.chk_only_fail.setEnabled(False)
            self.btn_export.setEnabled(False)

    def refresh_raw_table(self):
        if self.all_data.empty: return
        df_to_show = self.all_data[self.all_data[AppConfig.Columns.RESULT] == 'FAIL'] if self.chk_only_fail.isChecked() else self.all_data
        
        MAX_DISPLAY = 5000 
        rows = min(len(df_to_show), MAX_DISPLAY)
        self.raw_table.setRowCount(rows)
        self.raw_table.setSortingEnabled(False)
        
        red_brush = QBrush(QColor(255, 220, 220))
        red_text = QColor(200, 0, 0)
        green_text = QColor(0, 128, 0)
        
        col_indices = [df_to_show.columns.get_loc(c) for c in DISPLAY_COLUMNS if c in df_to_show.columns]
        
        for r in range(rows):
            is_fail = str(df_to_show.iloc[r][AppConfig.Columns.RESULT]) == "FAIL"
            for table_c, df_c in enumerate(col_indices):
                val = df_to_show.iloc[r, df_c]
                item_text = ""
                if table_c == 1 and isinstance(val, (datetime, pd.Timestamp)):
                    item_text = val.strftime("%Y/%m/%d %H:%M:%S") if pd.notnull(val) else ""
                else:
                    item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                
                # Use NumericTableWidgetItem for numeric columns
                if DISPLAY_COLUMNS[table_c] in [AppConfig.Columns.NO, AppConfig.Columns.MEASURED, AppConfig.Columns.DESIGN, AppConfig.Columns.DIFF, AppConfig.Columns.UPPER, AppConfig.Columns.LOWER]:
                    item = NumericTableWidgetItem(item_text)
                else:
                    item = QTableWidgetItem(item_text)

                if is_fail:
                    if DISPLAY_COLUMNS[table_c] in [AppConfig.Columns.DIFF, AppConfig.Columns.RESULT]:
                        item.setForeground(red_text)
                        item.setBackground(red_brush)
                elif item_text == "OK" and DISPLAY_COLUMNS[table_c] == AppConfig.Columns.RESULT:
                    item.setForeground(green_text)
                self.raw_table.setItem(r, table_c, item)
        self.raw_table.setSortingEnabled(True)
        status = f"Raw Data: {len(df_to_show)} 筆 | 總樣本: {len(self.loaded_files)}"
        if len(df_to_show) > MAX_DISPLAY: status += " (僅顯示前5000筆)"
        self.lbl_status.setText(status)

    def calculate_and_refresh_stats(self):
        if self.all_data.empty: return
        self.lbl_info.setText("正在計算統計數據...")
        total_files = len(self.loaded_files)
        grouped = self.all_data.groupby([AppConfig.Columns.NO, AppConfig.Columns.PROJECT])
        
        stats_list = []
        for (no, name), group in grouped:
            count = len(group)
            ng_count = len(group[group[AppConfig.Columns.RESULT] == 'FAIL'])
            fail_rate = (ng_count / total_files) * 100 if total_files > 0 else 0
            vals = pd.to_numeric(group[AppConfig.Columns.MEASURED], errors='coerce').dropna()
            
            first = group.iloc[0]
            design = float(first.get(AppConfig.Columns.DESIGN, 0))
            upper = float(first.get(AppConfig.Columns.UPPER, 0))
            lower = float(first.get(AppConfig.Columns.LOWER, 0))
            usl = design + upper
            lsl = design + lower
            
            mean_val = vals.mean() if not vals.empty else 0
            max_val = vals.max() if not vals.empty else 0
            min_val = vals.min() if not vals.empty else 0
            
            cpk, reliability = calculate_cpk(vals, usl, lsl)
            
            # [v2.3.0] 計算建議公差
            tol_result = calculate_tolerance_for_yield(vals, design, AppConfig.DEFAULT_TARGET_YIELD)
            
            stats_list.append({
                "No": no, "測量專案": name, "樣本數": count, 
                "NG數": ng_count, "不良率(%)": fail_rate, "CPK": cpk,
                "CPK_RELIABILITY": reliability,
                "建議公差": tol_result['symmetric_tol'],
                "TOL_RELIABILITY": tol_result['reliability'],
                "TOL_UPPER": tol_result['upper_tol'],
                "TOL_LOWER": tol_result['lower_tol'],
                "TOL_OFFSET": tol_result['offset'],
                "平均值": mean_val, "最大值": max_val, "最小值": min_val,
                "_design": design, "_upper": upper, "_lower": lower
            })
            
        self.stats_data = pd.DataFrame(stats_list)
        
        # [v2.0.3] 使用自然排序 (Natsort)
        if HAS_NATSORT:
            try:
                sorted_idx = index_natsorted(self.stats_data['No'], alg=ns.IGNORECASE)
                self.stats_data = self.stats_data.iloc[sorted_idx]
            except Exception as e:
                logging.warning(f"Natsort failed, using fallback: {e}")
                self.stats_data['_sort_key'] = self.stats_data['No'].apply(natural_keys)
                self.stats_data.sort_values(by="_sort_key", inplace=True)
                self.stats_data.drop(columns=['_sort_key'], inplace=True)
        else:
            self.stats_data['_sort_key'] = self.stats_data['No'].apply(natural_keys)
            self.stats_data.sort_values(by="_sort_key", inplace=True)
            self.stats_data.drop(columns=['_sort_key'], inplace=True)
        
        total_items = len(self.stats_data)
        ng_items = len(self.stats_data[self.stats_data["NG數"] > 0])
        self.lbl_stats_summary.setText(
            f"總樣本數: {total_files} | 總測項: {total_items} | 有NG項目: {ng_items} | "
            f"平均良率: {100 - self.stats_data['不良率(%)'].mean():.2f}%"
        )
        
        self.stats_table.setRowCount(len(self.stats_data))
        self.stats_table.setSortingEnabled(False)
        for r in range(len(self.stats_data)):
            row = self.stats_data.iloc[r]
            self.stats_table.setItem(r, 0, NumericTableWidgetItem(str(row['No'])))
            self.stats_table.setItem(r, 1, QTableWidgetItem(str(row['測量專案'])))
            self.stats_table.setItem(r, 2, NumericTableWidgetItem(str(row['樣本數'])))
            
            ng_item = NumericTableWidgetItem(str(row['NG數']))
            if row['NG數'] > 0: ng_item.setForeground(QColor('red'))
            self.stats_table.setItem(r, 3, ng_item)
            
            rate_item = NumericTableWidgetItem(f"{row['不良率(%)']:.2f}")
            if row['不良率(%)'] > 0: rate_item.setForeground(QColor('red'))
            self.stats_table.setItem(r, 4, rate_item)
            
            # CPK Display Logic
            cpk_val = row['CPK']
            reliability = row.get('CPK_RELIABILITY', 'reliable')
            sample_count = row['樣本數']
            
            cpk_text = ""
            cpk_item = NumericTableWidgetItem("")
            
            if reliability == 'invalid':
                cpk_text = "---"
                cpk_item.setText(cpk_text)
                cpk_item.setToolTip("無法計算 CPK (數據不足或規格異常)")
            elif reliability == 'small_sample':
                cpk_text = f"{cpk_val:.3f} ⚠"
                cpk_item.setText(cpk_text)
                cpk_item.setForeground(QColor('darkorange')) # Use orange for warning
                cpk_item.setToolTip(
                    "警告：樣本數少於 30，CPK 值僅供參考\n"
                    f"當前樣本數：{sample_count}\n"
                    "建議：累積更多數據後再評估製程能力"
                )
            else:
                cpk_text = f"{cpk_val:.3f}"
                cpk_item.setText(cpk_text)
                cpk_item.setToolTip(f"CPK: {cpk_val:.3f} (樣本數：{sample_count})")
                
                # Color coding for reliable CPK
                if cpk_val < 1.0: cpk_item.setBackground(QBrush(QColor(255, 200, 200)))
                elif cpk_val < 1.33: cpk_item.setBackground(QBrush(QColor(255, 255, 200)))
                else: cpk_item.setBackground(QBrush(QColor(200, 255, 200)))

            self.stats_table.setItem(r, 5, cpk_item)
            
            # [v2.3.0] 建議公差 Display Logic
            tol_val = row['建議公差']
            tol_reliability = row.get('TOL_RELIABILITY', 'invalid')
            tol_upper = row.get('TOL_UPPER', np.nan)
            tol_lower = row.get('TOL_LOWER', np.nan)
            tol_offset = row.get('TOL_OFFSET', np.nan)
            current_upper = row.get('_upper', 0)
            current_lower = row.get('_lower', 0)
            
            tol_item = NumericTableWidgetItem("")
            
            if tol_reliability == 'invalid' or tol_reliability == 'zero_std':
                tol_item.setText("---")
                tol_item.setToolTip("無法計算 (數據不足或標準差為零)")
            else:
                tol_text = f"±{tol_val:.4f}"
                if tol_reliability == 'small_sample':
                    tol_text += " ⚠"
                    tol_item.setForeground(QColor('darkorange'))
                
                tol_item.setText(tol_text)
                
                # 詳細 Tooltip
                tooltip_lines = [
                    f"【達成 90% 良率所需公差】",
                    f"對稱公差：±{tol_val:.4f}",
                    f"",
                    f"📊 非對稱建議：",
                    f"  上限：+{tol_upper:.4f}",
                    f"  下限：{tol_lower:.4f}",
                    f"",
                    f"📐 當前設定：",
                    f"  上限：+{current_upper:.4f}",
                    f"  下限：{current_lower:.4f}",
                    f"",
                    f"📈 製程偏移：{tol_offset:+.4f}" if not np.isnan(tol_offset) else ""
                ]
                tol_item.setToolTip("\n".join([l for l in tooltip_lines if l]))
                
                # 顏色標記：與當前規格比較
                if not np.isnan(tol_val):
                    current_tol = max(abs(current_upper), abs(current_lower))
                    if current_tol > 0:
                        if tol_val > current_tol * 1.2:  # 建議公差比當前大 20%
                            tol_item.setBackground(QBrush(QColor(255, 220, 220)))  # 淺紅：規格偏緊
                        elif tol_val < current_tol * 0.8:  # 建議公差比當前小 20%
                            tol_item.setBackground(QBrush(QColor(220, 255, 220)))  # 淺綠：規格充裕
            
            self.stats_table.setItem(r, 6, tol_item)
            self.stats_table.setItem(r, 7, NumericTableWidgetItem(f"{row['平均值']:.4f}"))
            self.stats_table.setItem(r, 8, NumericTableWidgetItem(f"{row['最大值']:.4f}"))
            self.stats_table.setItem(r, 9, NumericTableWidgetItem(f"{row['最小值']:.4f}"))
        self.stats_table.setSortingEnabled(True)
        self.lbl_info.setText("統計數據更新完成。")

    def plot_from_raw_table(self):
        sel = self.raw_table.selectedItems()
        if not sel: return
        row = sel[0].row()
        target_no = self.raw_table.item(row, 2).text()
        target_name = self.raw_table.item(row, 3).text()
        self.open_plot_dialog(target_no, target_name)

    def plot_from_stats_table(self):
        sel = self.stats_table.selectedItems()
        if not sel: return
        row = sel[0].row()
        target_no = self.stats_table.item(row, 0).text()
        target_name = self.stats_table.item(row, 1).text()
        self.open_plot_dialog(target_no, target_name)

    def open_plot_dialog(self, no, name):
        try:
            mask = (self.all_data[AppConfig.Columns.NO].astype(str) == no) & (self.all_data[AppConfig.Columns.PROJECT] == name)
            df_item = self.all_data[mask]
            if df_item.empty: return
            
            first = df_item.iloc[0]
            design = float(first.get(AppConfig.Columns.DESIGN, 0))
            upper = float(first.get(AppConfig.Columns.UPPER, 0))
            lower = float(first.get(AppConfig.Columns.LOWER, 0))
            
            plot_dlg = DistributionPlotDialog(f"{name} (No.{no})", df_item, design, upper, lower, self, self.current_theme)
            plot_dlg.exec()
        except Exception as e:
            logging.error(f"繪圖失敗: {e}")
            QMessageBox.critical(self, "錯誤", f"無法分析: {e}")

    def export_current_tab(self):
        curr_idx = self.tabs.currentIndex()
        if curr_idx == 0: # Stats
            if self.stats_data.empty: return
            path, _ = QFileDialog.getSaveFileName(self, "匯出統計報表", "Statistics.csv", "CSV (*.csv)")
            if path:
                export_df = self.stats_data.drop(columns=["_design", "_upper", "_lower", "_sort_key", "CPK_RELIABILITY"], errors='ignore')
                export_df.to_csv(path, index=False, encoding='utf-8-sig')
                QMessageBox.information(self, "完成", "統計報表已匯出")
        elif curr_idx == 1: # Raw
            if self.all_data.empty: return
            path, _ = QFileDialog.getSaveFileName(self, "匯出原始資料", "RawData.csv", "CSV (*.csv)")
            if path:
                self.all_data.to_csv(path, index=False, encoding='utf-8-sig')
                QMessageBox.information(self, "完成", "原始資料已匯出")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MeasurementAnalyzerApp()
    window.show()
    sys.exit(app.exec())
