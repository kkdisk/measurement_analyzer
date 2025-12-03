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
                             QTabWidget, QTextEdit, QSplitter, QFrame, QMenu, QInputDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QBrush, QFont, QAction

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

# --- 設定常數 ---
APP_VERSION = "v2.0.2"
APP_TITLE = f"量測數據分析工具 (Pro版) {APP_VERSION}"
LOG_FILENAME = "measurement_analyzer.log"
THEME_CONFIG_FILE = "theme_config.txt"

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
COL_ORIGINAL_JUDGE = '判斷'
COL_ORIGINAL_JUDGE_PDF = '判断'

DISPLAY_COLUMNS = [
    COL_FILE, COL_TIME, COL_NO, COL_PROJECT, 
    COL_MEASURED, COL_DESIGN, COL_DIFF, 
    COL_UPPER, COL_LOWER, COL_RESULT
]

UPDATE_LOG = """
=== 版本更新紀錄 ===

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
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILENAME, encoding='utf-8', mode='w'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    logging.getLogger("matplotlib").setLevel(logging.WARNING)
    
    logging.info(f"應用程式啟動 - {APP_TITLE}")

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
        
        rows = {} 
        for word in words:
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
                                COL_NO: d['no'],
                                COL_PROJECT: d['proj'].strip(),
                                COL_MEASURED: d['val'],
                                COL_UNIT: d['unit'],
                                COL_DESIGN: d['design'],
                                COL_UPPER: d['up'],
                                COL_LOWER: d['low'],
                                COL_ORIGINAL_JUDGE: d['judge']
                            }
                            data_list.append(item)
            
            if data_list:
                df = pd.DataFrame(data_list)
                return df, measure_time
            else:
                return None, None

        except Exception as e:
            logging.error(f"PDF Error {filepath}: {e}")
            return None, None

    def run(self):
        new_data_frames = []
        loaded_filenames = set()
        for i, filepath in enumerate(self.file_paths):
            if not self._is_running: break
            filename = os.path.basename(filepath)
            self.progress_updated.emit(i + 1, f"處理中: {filename}")
            
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
                    if COL_NO not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: COL_NO}, inplace=True)
                                break
                    required = [COL_NO, COL_MEASURED, COL_DESIGN]
                    if all(c in df.columns for c in required):
                        df = df.dropna(subset=[COL_NO])
                        num_cols = [COL_MEASURED, COL_DESIGN, COL_UPPER, COL_LOWER]
                        for c in num_cols:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors='coerce')
                            else:
                                df[c] = 0.0
                        
                        df[COL_DIFF] = df[COL_MEASURED] - df[COL_DESIGN]
                        df[COL_RESULT] = "OK"
                        
                        mask_ignore = df[COL_DESIGN].abs() < 0.000001
                        df.loc[mask_ignore, COL_RESULT] = "---"
                        
                        mask_tol_na = df[COL_UPPER].isna() | df[COL_LOWER].isna()
                        df.loc[mask_tol_na, COL_RESULT] = "---"
                        
                        mask_tol_zero = (df[COL_UPPER] == 0) & (df[COL_LOWER] == 0)
                        orig_judge = None
                        if COL_ORIGINAL_JUDGE in df.columns: orig_judge = COL_ORIGINAL_JUDGE
                        elif COL_ORIGINAL_JUDGE_PDF in df.columns: orig_judge = COL_ORIGINAL_JUDGE_PDF
                        
                        if orig_judge:
                            df.loc[mask_tol_zero, COL_RESULT] = df.loc[mask_tol_zero, orig_judge].fillna("---")
                        else:
                            df.loc[mask_tol_zero, COL_RESULT] = "---"
                            
                        mask_check = ~(mask_ignore | mask_tol_na | mask_tol_zero)
                        mask_fail = mask_check & ((df[COL_DIFF] > df[COL_UPPER]) | (df[COL_DIFF] < df[COL_LOWER]))
                        df.loc[mask_fail, COL_RESULT] = "FAIL"
                        
                        df[COL_FILE] = filename
                        df[COL_TIME] = measure_time if measure_time else pd.NaT
                        if COL_PROJECT not in df.columns: df[COL_PROJECT] = ''
                        
                        cols = [c for c in DISPLAY_COLUMNS if c in df.columns]
                        new_data_frames.append(df[cols])
            except Exception as e:
                logging.error(f"Error {filename}: {e}")

        self.data_loaded.emit(new_data_frames, loaded_filenames)

    def stop(self):
        self._is_running = False

# --- GUI Components ---
class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text()) < float(other.text())
        except ValueError:
            return super().__lt__(other)

class VersionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 600, 450)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(APP_TITLE))
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
        self.setGeometry(100, 100, 900, 600)
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
        if COL_TIME in df_sorted.columns:
            try:
                if len(df_sorted[COL_TIME].dropna()) > 0:
                    df_sorted = df_sorted.sort_values(by=COL_TIME)
                    has_time = True
            except: pass
        
        y_data = df_sorted[COL_MEASURED]
        x_data = range(1, len(y_data) + 1)
        
        line_color = 'cyan' if self.theme == 'dark' else 'blue'
        ax.plot(x_data, y_data, marker='o', linestyle='-', color=line_color, markersize=4, label='實測值')
        ax.axhline(self.design_val, color='lime' if self.theme=='dark' else 'green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        ax.set_title("量測值趨勢圖")
        ax.set_xlabel("時間順序" if has_time else "讀取順序")
        ax.legend()
        ax.grid(True, alpha=0.3)
        layout.addWidget(toolbar)
        layout.addWidget(canvas)

# --- 主程式 ---
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
        self.current_theme = 'light'
        self.init_theme()
        self.init_ui()

    def init_theme(self):
        if not HAS_THEME_SUPPORT: return
        try:
            if os.path.exists(THEME_CONFIG_FILE):
                with open(THEME_CONFIG_FILE, 'r') as f:
                    self.current_theme = f.read().strip()
            qdarktheme.setup_theme(self.current_theme)
        except Exception as e:
            logging.error(f"主題載入失敗: {e}")

    def toggle_theme(self):
        if not HAS_THEME_SUPPORT:
            QMessageBox.information(self, "提示", "請先安裝 'pyqtdarktheme' 套件")
            return
        new_theme = 'dark' if self.current_theme == 'light' else 'light'
        self.current_theme = new_theme
        qdarktheme.setup_theme(new_theme)
        try:
            with open(THEME_CONFIG_FILE, 'w') as f:
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
        
        self.btn_clear = QPushButton("清空資料")
        self.btn_clear.clicked.connect(self.clear_all_data)
        self.btn_clear.setStyleSheet("color: red;")
        
        self.btn_export = QPushButton("匯出當前頁面資料")
        self.btn_export.clicked.connect(self.export_current_tab)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)
        
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
        cols = ["No", "測量專案", "樣本數", "NG數", "不良率(%)", "CPK", "平均值", "最大值", "最小值"]
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
            
            msg = f"完成。本次加入 {len(new_data)} 筆數據。"
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
        df_to_show = self.all_data[self.all_data[COL_RESULT] == 'FAIL'] if self.chk_only_fail.isChecked() else self.all_data
        
        MAX_DISPLAY = 5000 
        rows = min(len(df_to_show), MAX_DISPLAY)
        self.raw_table.setRowCount(rows)
        self.raw_table.setSortingEnabled(False)
        
        red_brush = QBrush(QColor(255, 220, 220))
        red_text = QColor(200, 0, 0)
        green_text = QColor(0, 128, 0)
        
        col_indices = [df_to_show.columns.get_loc(c) for c in DISPLAY_COLUMNS if c in df_to_show.columns]
        
        for r in range(rows):
            is_fail = str(df_to_show.iloc[r][COL_RESULT]) == "FAIL"
            for table_c, df_c in enumerate(col_indices):
                val = df_to_show.iloc[r, df_c]
                item_text = ""
                if table_c == 1 and isinstance(val, (datetime, pd.Timestamp)):
                    item_text = val.strftime("%Y/%m/%d %H:%M:%S") if pd.notnull(val) else ""
                else:
                    item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                
                # Use NumericTableWidgetItem for numeric columns
                if DISPLAY_COLUMNS[table_c] in [COL_NO, COL_MEASURED, COL_DESIGN, COL_DIFF, COL_UPPER, COL_LOWER]:
                    item = NumericTableWidgetItem(item_text)
                else:
                    item = QTableWidgetItem(item_text)

                if is_fail:
                    if DISPLAY_COLUMNS[table_c] in [COL_DIFF, COL_RESULT]:
                        item.setForeground(red_text)
                        item.setBackground(red_brush)
                elif item_text == "OK" and DISPLAY_COLUMNS[table_c] == COL_RESULT:
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
        grouped = self.all_data.groupby([COL_NO, COL_PROJECT])
        
        stats_list = []
        for (no, name), group in grouped:
            count = len(group)
            ng_count = len(group[group[COL_RESULT] == 'FAIL'])
            fail_rate = (ng_count / total_files) * 100 if total_files > 0 else 0
            vals = pd.to_numeric(group[COL_MEASURED], errors='coerce').dropna()
            
            first = group.iloc[0]
            design = float(first.get(COL_DESIGN, 0))
            upper = float(first.get(COL_UPPER, 0))
            lower = float(first.get(COL_LOWER, 0))
            usl = design + upper
            lsl = design + lower
            
            mean_val = vals.mean() if not vals.empty else 0
            max_val = vals.max() if not vals.empty else 0
            min_val = vals.min() if not vals.empty else 0
            
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
                "_design": design, "_upper": upper, "_lower": lower
            })
            
        self.stats_data = pd.DataFrame(stats_list)
        self.stats_data['_sort_no'] = pd.to_numeric(self.stats_data['No'], errors='coerce')
        self.stats_data.sort_values(by=["不良率(%)", "_sort_no"], ascending=[False, True], inplace=True)
        
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
            
            cpk_val = row['CPK']
            cpk_text = "---" if pd.isna(cpk_val) else f"{cpk_val:.3f}"
            cpk_item = NumericTableWidgetItem(cpk_text)
            if not pd.isna(cpk_val):
                if row['樣本數'] < 30:
                    cpk_item.setForeground(QBrush(QColor('gray')))
                    cpk_item.setText(f"{cpk_text} (少)")
                else:
                    if cpk_val < 1.0: cpk_item.setBackground(QBrush(QColor(255, 200, 200)))
                    elif cpk_val < 1.33: cpk_item.setBackground(QBrush(QColor(255, 255, 200)))
                    else: cpk_item.setBackground(QBrush(QColor(200, 255, 200)))
            self.stats_table.setItem(r, 5, cpk_item)
            self.stats_table.setItem(r, 6, NumericTableWidgetItem(f"{row['平均值']:.4f}"))
            self.stats_table.setItem(r, 7, NumericTableWidgetItem(f"{row['最大值']:.4f}"))
            self.stats_table.setItem(r, 8, NumericTableWidgetItem(f"{row['最小值']:.4f}"))
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
        curr_idx = self.tabs.currentIndex()
        if curr_idx == 0: # Stats
            if self.stats_data.empty: return
            path, _ = QFileDialog.getSaveFileName(self, "匯出統計報表", "Statistics.csv", "CSV (*.csv)")
            if path:
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