import sys
import os
import glob
import pandas as pd
import numpy as np
import re
import logging
import traceback
from datetime import datetime

# PyQt6 imports
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QProgressBar, QMessageBox, QGroupBox, QCheckBox, QDialog,
                             QTabWidget, QTextEdit, QSplitter, QFrame, QMenu)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QBrush, QFont, QAction
import qdarktheme

# Matplotlib imports for plotting
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt

# --- 設定常數與欄位名稱 (方便統一管理) ---
APP_VERSION = "v1.7.0"
APP_TITLE = f"量測數據分析工具 (Pro版) {APP_VERSION}"
LOG_FILENAME = "measurement_analyzer.log"
THEME_CONFIG = "theme_config.txt"

# 欄位定義
COL_FILE = '檔案名稱'
COL_TIME = '測量時間'
COL_NO = 'No'
COL_PROJECT = '測量專案'
COL_MEASURED = '實測值'
COL_DESIGN = '設計值'
COL_DIFF = '差異'
COL_UPPER = '上限公差'
COL_LOWER = '下限公差'
COL_RESULT = '判定結果'
COL_UNIT = '單位'
COL_ORIGINAL_JUDGE = '判斷' # 原檔的判斷欄位

# 顯示順序
DISPLAY_COLUMNS = [
    COL_FILE, COL_TIME, COL_NO, COL_PROJECT, 
    COL_MEASURED, COL_DESIGN, COL_DIFF, 
    COL_UPPER, COL_LOWER, COL_RESULT
]

UPDATE_LOG = """
=== 版本更新紀錄 ===

[v1.7.0] - 2025/12/03 (UI Modernization)
1. [外觀] 新增「深色模式 (Dark Mode)」與「淺色模式」，可於選單列切換。
2. [介面] 全面優化 UI 元件風格 (圓角按鈕、現代化表格)。
3. [圖表] 視覺化圖表現在會自動適應深色/淺色背景。

[v1.6.1] - 2025/12/03 (UX Optimization)
1. [介面] 調整分頁順序：將「統計摘要分析」設為主頁 (Tab 1)，「原始數據」為次頁 (Tab 2)。
2. [排序] 修正統計表排序邏輯：No 欄位現在會依照數值大小 (1, 2, 3... 10) 排序，而非文字順序 (1, 10, 11...)。

[v1.6.0] - 2025/12/03
1. [介面] 導入分頁系統 (Tabs)：原始數據與統計摘要分頁顯示。
2. [互動] 統計頁面支援點擊測項直接開啟視覺化圖表。
3. [效能] 統計運算優化。
"""

# --- 初始化日誌系統 ---
def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILENAME, encoding='utf-8', mode='w'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info(f"應用程式啟動 - {APP_TITLE}")

# --- 設定中文字型 ---
import matplotlib.font_manager as fm
def get_chinese_font():
    font_names = ['Microsoft JhengHei', 'SimHei', 'PingFang TC', 'Arial Unicode MS']
    for name in font_names:
        if name in [f.name for f in fm.fontManager.ttflist]:
            return name
    return None

def set_chinese_font():
    font_name = get_chinese_font()
    if font_name:
        matplotlib.rcParams['font.sans-serif'] = [font_name]
        matplotlib.rcParams['axes.unicode_minus'] = False 
set_chinese_font()

# --- 日期解析工具 ---
def parse_keyence_date(date_str):
    if not isinstance(date_str, str): return None
    date_str = date_str.strip()
    try:
        match = re.match(r'(\d+)/(\d+)/(\d+)\s+(上午|下午)\s+(\d+):(\d+):(\d+)', date_str)
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
    except Exception as e:
        return None

# --- 背景工作執行緒 ---
class FileLoaderThread(QThread):
    progress_updated = pyqtSignal(int, str)
    data_loaded = pyqtSignal(list, set)
    error_occurred = pyqtSignal(str)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._is_running = True

    def find_header_row_and_date(self, filepath):
        try:
            encodings = ['utf-8-sig', 'big5', 'cp950', 'shift_jis']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        # 只讀前 60 行來判斷，避免讀取整個大檔案
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
                except StopIteration: continue
                except UnicodeDecodeError: continue 
            return None, None, None
        except Exception as e:
            logging.error(f"檔頭掃描錯誤 {filepath}: {e}")
            return None, None, None

    def run(self):
        new_data_frames = []
        loaded_filenames = set()
        total = len(self.file_paths)
        
        logging.info(f"Thread 開始處理 {total} 個檔案")

        for i, filepath in enumerate(self.file_paths):
            if not self._is_running: break 

            filename = os.path.basename(filepath)
            self.progress_updated.emit(i + 1, f"處理中: {filename}")
            
            try:
                header_idx, encoding, measure_time = self.find_header_row_and_date(filepath)
                
                if header_idx is not None:
                    loaded_filenames.add(filename)
                    # 讀取 CSV
                    df = pd.read_csv(filepath, skiprows=header_idx, header=0, 
                                   encoding=encoding, on_bad_lines='skip', index_col=False)
                    
                    # 欄位正規化
                    df.columns = [str(c).strip() for c in df.columns]
                    if COL_NO not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: COL_NO}, inplace=True)
                                break
                    
                    required = [COL_NO, COL_MEASURED, COL_DESIGN]
                    if all(c in df.columns for c in required):
                        df = df.dropna(subset=[COL_NO])
                        
                        # 轉數值
                        num_cols = [COL_MEASURED, COL_DESIGN, COL_UPPER, COL_LOWER]
                        for c in num_cols:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors='coerce')
                            else:
                                df[c] = 0.0
                        
                        # 向量化運算 (Vectorization) - 效能關鍵
                        df[COL_DIFF] = df[COL_MEASURED] - df[COL_DESIGN]
                        
                        # 預設 OK
                        df[COL_RESULT] = "OK"
                        
                        # 條件 1: 設計值太小 (視為參考值) -> ---
                        mask_ignore = df[COL_DESIGN].abs() < 0.000001
                        df.loc[mask_ignore, COL_RESULT] = "---"
                        
                        # 條件 2: 公差無效 -> ---
                        mask_tol_na = df[COL_UPPER].isna() | df[COL_LOWER].isna()
                        df.loc[mask_tol_na, COL_RESULT] = "---"
                        
                        # 條件 3: 公差為 0 (平面度等) -> 看原檔判斷
                        mask_tol_zero = (df[COL_UPPER] == 0) & (df[COL_LOWER] == 0)
                        if COL_ORIGINAL_JUDGE in df.columns:
                            # 填入原檔判斷，若原檔空則 ---
                            df.loc[mask_tol_zero, COL_RESULT] = df.loc[mask_tol_zero, COL_ORIGINAL_JUDGE].fillna("---")
                        else:
                            df.loc[mask_tol_zero, COL_RESULT] = "---"
                        
                        # 條件 4: 超規判定 (排除上述條件後)
                        # 需先排除 ignore, tol_na, tol_zero 這些行，避免覆蓋
                        mask_valid_check = ~(mask_ignore | mask_tol_na | mask_tol_zero)
                        mask_fail = mask_valid_check & (
                            (df[COL_DIFF] > df[COL_UPPER]) | 
                            (df[COL_DIFF] < df[COL_LOWER])
                        )
                        df.loc[mask_fail, COL_RESULT] = "FAIL"

                        # Metadata
                        df[COL_FILE] = filename
                        df[COL_TIME] = measure_time if measure_time else pd.NaT
                        if COL_PROJECT not in df.columns: df[COL_PROJECT] = ''

                        # 篩選欄位
                        cols = [c for c in DISPLAY_COLUMNS if c in df.columns]
                        new_data_frames.append(df[cols])
            
            except Exception as e:
                logging.error(f"檔案處理失敗 {filename}: {e}")

        logging.info(f"Thread 完成。成功: {len(loaded_filenames)}")
        self.data_loaded.emit(new_data_frames, loaded_filenames)

    def stop(self):
        self._is_running = False

# --- 對話視窗 ---
class VersionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 600, 450)
        layout = QVBoxLayout(self)
        lbl = QLabel(APP_TITLE)
        lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: darkblue;")
        layout.addWidget(lbl)
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(UPDATE_LOG)
        txt.setStyleSheet("font-family: Consolas; font-size: 12px;")
        layout.addWidget(txt)
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)

class DistributionPlotDialog(QDialog):
    def __init__(self, item_name, df_item, design_val, upper_tol, lower_tol, parent=None, theme='dark'):
        super().__init__(parent)
        self.setWindowTitle(f"詳細分析: {item_name}")
        self.setGeometry(100, 100, 900, 600)
        self.df_item = df_item
        self.design_val = design_val
        self.upper_tol = upper_tol
        self.lower_tol = lower_tol
        self.usl = design_val + upper_tol
        self.lsl = design_val + lower_tol
        self.current_theme = theme
        
        # 設定 Matplotlib 風格
        if self.current_theme == 'dark':
            plt.style.use('dark_background')
        else:
            plt.style.use('default')
            
        # [Bug Fix] 設定風格後必須重新套用中文字型，否則會變框框
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
        
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)

    def plot_histogram(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        data = self.df_item[COL_MEASURED].dropna()
        if len(data) > 0:
            ax.hist(data, bins=15, color='skyblue', edgecolor='black', alpha=0.7, label='實測值')
            ax.axvline(self.design_val, color='green', linestyle='-', linewidth=2, label='設計值')
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
        if COL_TIME in df_sorted.columns:
            try:
                if len(df_sorted[COL_TIME].dropna()) > 0:
                    df_sorted = df_sorted.sort_values(by=COL_TIME)
                    has_time = True
            except: pass
        
        y_data = df_sorted[COL_MEASURED]
        x_data = range(1, len(y_data) + 1)
        
        ax.plot(x_data, y_data, marker='o', linestyle='-', color='blue', markersize=4, label='實測值')
        ax.axhline(self.design_val, color='green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        ax.set_title("量測值趨勢圖")
        ax.set_xlabel("時間順序" if has_time else "讀取順序")
        ax.legend()
        ax.grid(True, alpha=0.3)
        layout.addWidget(toolbar)
        layout.addWidget(canvas)


# --- 主應用程式 ---
class MeasurementAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        setup_logging()
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 1300, 850)
        self.all_data = pd.DataFrame()
        self.stats_data = pd.DataFrame()
        self.loaded_files = set()
        self.loader_thread = None 
        self.current_theme = "auto"
        self.load_theme_config()
        self.init_ui()
        self.apply_theme(self.current_theme)

    def load_theme_config(self):
        if os.path.exists(THEME_CONFIG):
            try:
                with open(THEME_CONFIG, "r") as f:
                    self.current_theme = f.read().strip()
            except: pass

    def save_theme_config(self, theme):
        try:
            with open(THEME_CONFIG, "w") as f:
                f.write(theme)
        except: pass

    def apply_theme(self, theme):
        self.current_theme = theme
        qdarktheme.setup_theme(theme)
        self.save_theme_config(theme)
        
        # 定義顏色變數
        if theme == 'dark':
            self.lbl_status.setStyleSheet("color: #4da6ff; font-weight: bold;") # 淺藍
            self.lbl_stats_summary.setStyleSheet("background-color: #2d2d2d; padding: 10px; font-weight: bold; color: white;")
            # Dark Mode 表格樣式: 選取時使用半透明背景，讓原本的顏色透出來
            table_style = """
                QTableWidget {
                    selection-background-color: rgba(255, 255, 255, 0.15);
                    selection-color: white;
                }
            """
        else:
            self.lbl_status.setStyleSheet("color: blue; font-weight: bold;")
            self.lbl_stats_summary.setStyleSheet("background-color: #f0f0f0; padding: 10px; font-weight: bold; color: black;")
            # Light Mode 表格樣式
            table_style = """
                QTableWidget {
                    selection-background-color: rgba(0, 0, 0, 0.1);
                    selection-color: black;
                }
            """
            
        self.raw_table.setStyleSheet(table_style)
        self.stats_table.setStyleSheet(table_style)
        
        # 重新整理表格以套用新的顏色邏輯
        self.refresh_raw_table()
        self.calculate_and_refresh_stats()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- 選單列 (Menu Bar) ---
        menubar = self.menuBar()
        view_menu = menubar.addMenu("檢視 (View)")
        
        theme_menu = view_menu.addMenu("主題 (Theme)")
        action_auto = QAction("自動 (Auto)", self)
        action_auto.triggered.connect(lambda: self.apply_theme("auto"))
        theme_menu.addAction(action_auto)
        
        action_dark = QAction("深色 (Dark)", self)
        action_dark.triggered.connect(lambda: self.apply_theme("dark"))
        theme_menu.addAction(action_dark)
        
        action_light = QAction("淺色 (Light)", self)
        action_light.triggered.connect(lambda: self.apply_theme("light"))
        theme_menu.addAction(action_light)

        # 1. 頂部控制區
        control_group = QGroupBox("操作控制")
        control_layout = QHBoxLayout()
        
        self.btn_add = QPushButton("1. 加入資料夾")
        self.btn_add.clicked.connect(self.add_folder_data)
        self.btn_add.setMinimumHeight(40)
        
        self.btn_clear = QPushButton("清空資料")
        self.btn_clear.clicked.connect(self.clear_all_data)
        self.btn_clear.setStyleSheet("color: red;")
        
        self.btn_export = QPushButton("匯出當前頁面資料")
        self.btn_export.clicked.connect(self.export_current_tab)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)
        
        self.btn_version = QPushButton("關於")
        self.btn_version.clicked.connect(self.show_version_info)
        
        control_layout.addWidget(self.btn_add, 1)
        control_layout.addWidget(self.btn_clear)
        control_layout.addWidget(self.btn_export, 1)
        control_layout.addWidget(self.btn_version)
        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        # 2. 中間分頁區 (TabWidget)
        self.tabs = QTabWidget()
        
        # 分頁 1: 統計數據 (改為第一順位)
        self.tab_stats = QWidget()
        self.setup_statistics_tab()
        self.tabs.addTab(self.tab_stats, "1. 統計摘要分析")
        
        # 分頁 2: 原始數據 (改為第二順位)
        self.tab_raw = QWidget()
        self.setup_raw_data_tab()
        self.tabs.addTab(self.tab_raw, "2. 原始數據列表")
        
        main_layout.addWidget(self.tabs)

        # 3. 底部狀態列
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
        
        # 篩選器
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

        # 表格
        self.raw_table = QTableWidget()
        self.raw_table.setColumnCount(len(DISPLAY_COLUMNS))
        self.raw_table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.raw_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.raw_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.raw_table.setAlternatingRowColors(True)
        
        header = self.raw_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(self.raw_table)

    def setup_statistics_tab(self):
        layout = QVBoxLayout(self.tab_stats)
        
        # 摘要資訊區
        self.lbl_stats_summary = QLabel("尚未載入資料")
        self.lbl_stats_summary.setStyleSheet("background-color: #f0f0f0; padding: 10px; font-weight: bold;")
        layout.addWidget(self.lbl_stats_summary)
        
        # 提示
        lbl_hint = QLabel("提示：雙擊表格任一列可開啟詳細圖表分析")
        lbl_hint.setStyleSheet("color: gray; font-style: italic;")
        layout.addWidget(lbl_hint)

        # 統計表格
        self.stats_table = QTableWidget()
        cols = ["No", "測量專案", "樣本數", "NG數", "不良率(%)", "CPK", "平均值", "最大值", "最小值"]
        self.stats_table.setColumnCount(len(cols))
        self.stats_table.setHorizontalHeaderLabels(cols)
        self.stats_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.stats_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.stats_table.setAlternatingRowColors(True)
        self.stats_table.setSortingEnabled(True)
        self.stats_table.doubleClicked.connect(self.plot_from_stats_table) # 雙擊事件
        
        header = self.stats_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        layout.addWidget(self.stats_table)

    def show_version_info(self):
        dlg = VersionDialog(self)
        dlg.exec()

    def add_folder_data(self):
        folder_path = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if not folder_path: return
        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        if not csv_files:
            QMessageBox.warning(self, "提示", "無 CSV 檔案。")
            return

        self.set_ui_loading_state(True)
        self.lbl_info.setText(f"開始處理: {len(csv_files)} 個檔案...")
        self.progress_bar.setMaximum(len(csv_files))
        self.progress_bar.setValue(0)

        self.loader_thread = FileLoaderThread(csv_files)
        self.loader_thread.progress_updated.connect(self.on_progress_updated)
        self.loader_thread.data_loaded.connect(self.on_data_loaded)
        self.loader_thread.start()

    def set_ui_loading_state(self, is_loading):
        self.btn_add.setEnabled(not is_loading)
        self.btn_clear.setEnabled(not is_loading)
        if is_loading:
            self.btn_export.setEnabled(False)

    def on_progress_updated(self, value, message):
        self.progress_bar.setValue(value)
        self.lbl_info.setText(message)

    def on_data_loaded(self, new_data_frames, loaded_filenames):
        self.loaded_files.update(loaded_filenames)
        
        if new_data_frames:
            self.lbl_info.setText("正在合併資料...")
            QApplication.processEvents() 
            
            new_data = pd.concat(new_data_frames, ignore_index=True)
            
            if self.all_data.empty:
                self.all_data = new_data
            else:
                self.all_data = pd.concat([self.all_data, new_data], ignore_index=True)
            
            # 更新介面
            self.btn_export.setEnabled(True)
            self.chk_only_fail.setEnabled(True)
            self.btn_plot_raw.setEnabled(True)
            
            # 更新兩個分頁的資料
            self.refresh_raw_table()
            self.calculate_and_refresh_stats()
            
            logging.info(f"載入完成，總計 {len(self.all_data)} 筆")
            self.lbl_info.setText(f"完成。本次加入 {len(new_data)} 筆數據。")
            QMessageBox.information(self, "完成", f"已加入 {len(loaded_filenames)} 個檔案。")
        else:
            self.lbl_info.setText("無有效數據。")
            QMessageBox.warning(self, "結果", "未提取到有效數據。")
            
        self.set_ui_loading_state(False)

    def clear_all_data(self):
        reply = QMessageBox.question(self, '確認', '確定清空所有資料？', 
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
        
        # 篩選邏輯
        if self.chk_only_fail.isChecked():
            df_to_show = self.all_data[self.all_data[COL_RESULT] == 'FAIL']
        else:
            df_to_show = self.all_data
            
        # 顯示資料 (限制前 5000 筆以防卡頓)
        MAX_DISPLAY = 5000 
        rows = min(len(df_to_show), MAX_DISPLAY)
        self.raw_table.setRowCount(rows)
        self.raw_table.setSortingEnabled(False)
        
        # 根據主題設定顏色
        if self.current_theme == 'dark':
            # Dark Mode: 深紅底，亮紅字
            fail_bg = QBrush(QColor(80, 0, 0)) 
            fail_fg = QColor(255, 100, 100)
            ok_fg = QColor(100, 255, 100)
        else:
            # Light Mode: 淺紅底，深紅字
            fail_bg = QBrush(QColor(255, 220, 220))
            fail_fg = QColor(200, 0, 0)
            ok_fg = QColor(0, 128, 0)
        
        # 預先取得欄位索引以加速迴圈
        col_indices = [df_to_show.columns.get_loc(c) for c in DISPLAY_COLUMNS if c in df_to_show.columns]
        
        for r in range(rows):
            is_fail = str(df_to_show.iloc[r][COL_RESULT]) == "FAIL"
            
            for table_c, df_c in enumerate(col_indices):
                val = df_to_show.iloc[r, df_c]
                item_text = ""
                
                if table_c == 1 and isinstance(val, (datetime, pd.Timestamp)): # Time
                    item_text = val.strftime("%Y/%m/%d %H:%M:%S") if pd.notnull(val) else ""
                else:
                    item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                
                item = QTableWidgetItem(item_text)
                
                if is_fail:
                    # 差異 & 判定欄位標紅
                    if DISPLAY_COLUMNS[table_c] in [COL_DIFF, COL_RESULT]:
                        item.setForeground(fail_fg)
                        item.setBackground(fail_bg)
                elif item_text == "OK" and DISPLAY_COLUMNS[table_c] == COL_RESULT:
                    item.setForeground(ok_fg)
                
                self.raw_table.setItem(r, table_c, item)
        
        self.raw_table.setSortingEnabled(True)
        
        status = f"Raw Data: {len(df_to_show)} 筆 | 總樣本: {len(self.loaded_files)}"
        if len(df_to_show) > MAX_DISPLAY: status += " (僅顯示前5000筆)"
        self.lbl_status.setText(status)

    def calculate_and_refresh_stats(self):
        """計算統計數據並更新分頁 2"""
        if self.all_data.empty: return
        
        self.lbl_info.setText("正在計算統計數據...")
        total_files = len(self.loaded_files)
        
        # GroupBy 運算
        grouped = self.all_data.groupby([COL_NO, COL_PROJECT])
        
        stats_list = []
        for (no, name), group in grouped:
            count = len(group)
            ng_count = len(group[group[COL_RESULT] == 'FAIL'])
            # 不良率基數: 總樣本數 (除非該測項在某些檔案未測到，此邏輯假設每片都有測)
            fail_rate = (ng_count / total_files) * 100 if total_files > 0 else 0
            
            vals = pd.to_numeric(group[COL_MEASURED], errors='coerce').dropna()
            
            # 規格
            first = group.iloc[0]
            design = float(first.get(COL_DESIGN, 0))
            upper = float(first.get(COL_UPPER, 0))
            lower = float(first.get(COL_LOWER, 0))
            usl = design + upper
            lsl = design + lower
            
            # 統計值
            mean_val = vals.mean() if not vals.empty else 0
            max_val = vals.max() if not vals.empty else 0
            min_val = vals.min() if not vals.empty else 0
            
            # CPK
            cpk = np.nan
            if len(vals) >= 2 and (usl != lsl):
                std = vals.std()
                if std == 0: cpk = 999.0
                else:
                    cpu = (usl - mean_val) / (3 * std)
                    cpl = (mean_val - lsl) / (3 * std)
                    cpk = min(cpu, cpl)
            
            stats_list.append({
                "No": no, "測量專案": name, "樣本數": count, 
                "NG數": ng_count, "不良率(%)": fail_rate, "CPK": cpk,
                "平均值": mean_val, "最大值": max_val, "最小值": min_val,
                # 隱藏欄位供繪圖用
                "_design": design, "_upper": upper, "_lower": lower
            })
            
        self.stats_data = pd.DataFrame(stats_list)
        
        # [排序邏輯修正] 
        # 1. 建立臨時排序欄位，將 No 轉為數字 (處理 1, 10, 2 字串排序問題)
        self.stats_data['_sort_no'] = pd.to_numeric(self.stats_data['No'], errors='coerce')
        
        # 2. 排序：不良率(高->低) > No(小->大)
        self.stats_data.sort_values(by=["不良率(%)", "_sort_no"], ascending=[False, True], inplace=True)
        
        # 更新 Summary Label
        total_items = len(self.stats_data)
        ng_items = len(self.stats_data[self.stats_data["NG數"] > 0])
        self.lbl_stats_summary.setText(
            f"總樣本數: {total_files} | 總測項: {total_items} | 有NG項目: {ng_items} | "
            f"平均良率: {100 - self.stats_data['不良率(%)'].mean():.2f}%"
        )
        
        # 填入表格
        self.stats_table.setRowCount(len(self.stats_data))
        self.stats_table.setSortingEnabled(False)
        
        # 根據主題設定顏色
        if self.current_theme == 'dark':
            fail_fg = QColor(255, 100, 100)
            rate_fg = QColor(255, 100, 100)
        else:
            fail_fg = QColor('red')
            rate_fg = QColor('red')

        for r in range(len(self.stats_data)):
            row = self.stats_data.iloc[r]
            
            self.stats_table.setItem(r, 0, QTableWidgetItem(str(row['No'])))
            self.stats_table.setItem(r, 1, QTableWidgetItem(str(row['測量專案'])))
            self.stats_table.setItem(r, 2, QTableWidgetItem(str(row['樣本數'])))
            
            ng_item = QTableWidgetItem(str(row['NG數']))
            if row['NG數'] > 0: ng_item.setForeground(fail_fg)
            self.stats_table.setItem(r, 3, ng_item)
            
            rate_item = QTableWidgetItem(f"{row['不良率(%)']:.2f}")
            if row['不良率(%)'] > 0: rate_item.setForeground(rate_fg)
            self.stats_table.setItem(r, 4, rate_item)
            
            # CPK Color
            cpk_val = row['CPK']
            cpk_text = "---" if pd.isna(cpk_val) else f"{cpk_val:.3f}"
            cpk_item = QTableWidgetItem(cpk_text)
            
            if not pd.isna(cpk_val):
                if row['樣本數'] < 30:
                    cpk_item.setForeground(QBrush(QColor('gray')))
                    cpk_item.setText(f"{cpk_text} (少)")
                else:
                    # CPK 顏色邏輯 (需適配 Dark Mode，這裡暫時維持背景色，但需注意對比)
                    if cpk_val < 1.0: 
                        cpk_item.setBackground(QBrush(QColor(255, 200, 200) if self.current_theme != 'dark' else QColor(100, 0, 0)))
                    elif cpk_val < 1.33: 
                        cpk_item.setBackground(QBrush(QColor(255, 255, 200) if self.current_theme != 'dark' else QColor(100, 100, 0)))
                    else: 
                        cpk_item.setBackground(QBrush(QColor(200, 255, 200) if self.current_theme != 'dark' else QColor(0, 80, 0)))
            self.stats_table.setItem(r, 5, cpk_item)
            
            self.stats_table.setItem(r, 6, QTableWidgetItem(f"{row['平均值']:.4f}"))
            self.stats_table.setItem(r, 7, QTableWidgetItem(f"{row['最大值']:.4f}"))
            self.stats_table.setItem(r, 8, QTableWidgetItem(f"{row['最小值']:.4f}"))
            
        self.stats_table.setSortingEnabled(True)
        self.lbl_info.setText("統計數據更新完成。")

    def plot_from_raw_table(self):
        """從原始資料表繪圖"""
        sel = self.raw_table.selectedItems()
        if not sel: return
        row = sel[0].row()
        target_no = self.raw_table.item(row, 2).text()
        target_name = self.raw_table.item(row, 3).text()
        self.open_plot_dialog(target_no, target_name)

    def plot_from_stats_table(self):
        """從統計表雙擊繪圖"""
        sel = self.stats_table.selectedItems()
        if not sel: return
        row = sel[0].row()
        target_no = self.stats_table.item(row, 0).text()
        target_name = self.stats_table.item(row, 1).text()
        self.open_plot_dialog(target_no, target_name)

    def open_plot_dialog(self, no, name):
        try:
            mask = (self.all_data[COL_NO].astype(str) == no) & (self.all_data[COL_PROJECT] == name)
            df_item = self.all_data[mask]
            if df_item.empty: return
            
            first = df_item.iloc[0]
            design = float(first.get(COL_DESIGN, 0))
            upper = float(first.get(COL_UPPER, 0))
            lower = float(first.get(COL_LOWER, 0))
            
            plot_dlg = DistributionPlotDialog(f"{name} (No.{no})", df_item, design, upper, lower, self, self.current_theme)
            plot_dlg.exec()
        except Exception as e:
            logging.error(f"繪圖失敗: {e}")
            QMessageBox.critical(self, "錯誤", f"無法分析: {e}")

    def export_current_tab(self):
        """匯出目前當前分頁的資料"""
        curr_idx = self.tabs.currentIndex()
        # 注意: 因為調整了分頁順序，現在 0 是 Stats, 1 是 Raw
        if curr_idx == 0: # Stats
            if self.stats_data.empty: return
            path, _ = QFileDialog.getSaveFileName(self, "匯出統計報表", "Statistics.csv", "CSV (*.csv)")
            if path:
                # 排除隱藏欄位
                export_df = self.stats_data.drop(columns=["_design", "_upper", "_lower", "_sort_no"], errors='ignore')
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