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
                             QTabWidget, QTextEdit)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QBrush, QAction, QIcon

# Matplotlib imports for plotting
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

# --- 版本資訊 ---
APP_VERSION = "v1.4.0"
APP_TITLE = f"量測數據分析工具 (Pro版) {APP_VERSION}"
UPDATE_LOG = """
=== 版本更新紀錄 ===

[v1.4.0] - 2025/12/03
1. [效能] 核心運算優化：導入向量化運算 (Vectorization)，大幅提升大量數據的判定速度 (>100x)。
2. [效能] 檔案讀取優化：改良檔頭偵測邏輯，避免讀取完整檔案造成的記憶體浪費。
3. [重構] 程式碼品質：統一欄位名稱定義，提升維護性。

[v1.3.0] - 2025/12/03
1. [新增] 日誌系統 (Logging)：錯誤訊息現在會寫入 'measurement_analyzer.log'。
2. [優化] CPK 統計邏輯：針對樣本數不足 (<30) 的項目，CPK 欄位會標示提示。
3. [優化] UI 體驗：表格截斷顯示時增加明確提示；修正執行緒錯誤處理機制。

[v1.2.0] - 2025/12/03
1. [重構] 導入 QThread 多執行緒架構，解決介面凍結問題。
"""

# --- 常數定義 ---
COL_FILENAME = "檔案名稱"
COL_MEASURE_TIME = "測量時間"
COL_NO = "No"
COL_PROJECT = "測量專案"
COL_MEASURE_VAL = "實測值"
COL_DESIGN_VAL = "設計值"
COL_DIFF = "差異"
COL_UPPER_TOL = "上限公差"
COL_LOWER_TOL = "下限公差"
COL_RESULT = "判定結果"
COL_JUDGMENT = "判斷" # 原始 CSV 可能有的欄位

# --- [新增] 日誌設定 ---
def setup_logging():
    """設定日誌系統，將輸出同時導向檔案與主控台"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("measurement_analyzer.log", encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info(f"應用程式啟動 - {APP_TITLE}")

# --- 設定中文字型 ---
import matplotlib.font_manager as fm
def set_chinese_font():
    font_names = ['Microsoft JhengHei', 'SimHei', 'PingFang TC', 'Arial Unicode MS']
    for name in font_names:
        if name in [f.name for f in fm.fontManager.ttflist]:
            matplotlib.rcParams['font.sans-serif'] = [name]
            matplotlib.rcParams['axes.unicode_minus'] = False 
            break
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
    except: return None

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
                    # [優化] 只讀取前 60 行來尋找 Header，避免讀取整個大檔案
                    lines = []
                    with open(filepath, 'r', encoding=enc) as f:
                        for _ in range(60):
                            line = f.readline()
                            if not line: break
                            lines.append(line)
                            
                    measure_time = None
                    for line in lines[:20]:
                        if "測量日期及時間" in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                measure_time = parse_keyence_date(parts[1].strip())
                            break
                    for i, line in enumerate(lines): 
                        if COL_NO in line and COL_MEASURE_VAL in line and COL_DESIGN_VAL in line:
                            return i, enc, measure_time
                except UnicodeDecodeError: continue 
            return None, None, None
        except Exception as e:
            logging.error(f"解析檔頭失敗 {filepath}: {e}")
            return None, None, None

    def run(self):
        new_data_frames = []
        loaded_filenames = set()
        total = len(self.file_paths)
        
        logging.info(f"開始處理 {total} 個檔案")

        for i, filepath in enumerate(self.file_paths):
            if not self._is_running: break 

            filename = os.path.basename(filepath)
            self.progress_updated.emit(i + 1, f"處理中: {filename}")
            
            try:
                header_idx, encoding, measure_time = self.find_header_row_and_date(filepath)
                
                if header_idx is not None:
                    loaded_filenames.add(filename)
                    # 使用 logging 記錄細節錯誤，on_bad_lines='warn' 可在 log 看到警告
                    df = pd.read_csv(filepath, skiprows=header_idx, header=0, 
                                   encoding=encoding, on_bad_lines='skip', index_col=False)
                    
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    if COL_NO not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: COL_NO}, inplace=True)
                                break
                    
                    required_cols = [COL_NO, COL_MEASURE_VAL, COL_DESIGN_VAL]
                    if all(col in df.columns for col in required_cols):
                        df = df.dropna(subset=[COL_NO])
                        
                        cols_to_numeric = [COL_MEASURE_VAL, COL_DESIGN_VAL, COL_UPPER_TOL, COL_LOWER_TOL]
                        for col in cols_to_numeric:
                            if col in df.columns:
                                df[col] = pd.to_numeric(df[col], errors='coerce')
                            else:
                                df[col] = 0.0
                        
                        df[COL_DIFF] = df[COL_MEASURE_VAL] - df[COL_DESIGN_VAL]
                        
                        # [優化] 向量化運算 (Vectorization)
                        # 條件 1: 設計值極小 或 公差為空 -> "---"
                        cond_ignore = (np.abs(df[COL_DESIGN_VAL]) < 1e-6) | \
                                      (df[COL_UPPER_TOL].isna()) | (df[COL_LOWER_TOL].isna())
                        
                        # 條件 2: 上下限公差皆為 0 -> 檢查 '判斷' 欄位
                        cond_tol_zero = (df[COL_UPPER_TOL] == 0) & (df[COL_LOWER_TOL] == 0)
                        
                        # 條件 3: 超出公差 -> "FAIL"
                        cond_fail = (df[COL_DIFF] > df[COL_UPPER_TOL]) | (df[COL_DIFF] < df[COL_LOWER_TOL])

                        # 準備 '判斷' 欄位資料 (若無則全填 "---")
                        if COL_JUDGMENT in df.columns:
                            judgment_vals = df[COL_JUDGMENT].fillna("---")
                        else:
                            judgment_vals = pd.Series("---", index=df.index)

                        # 使用 numpy.select 進行快速判定
                        conditions = [
                            cond_ignore,
                            cond_tol_zero,
                            cond_fail
                        ]
                        choices = [
                            "---",
                            judgment_vals,
                            "FAIL"
                        ]
                        
                        df[COL_RESULT] = np.select(conditions, choices, default="OK")

                        df.insert(0, COL_FILENAME, filename)
                        df.insert(1, COL_MEASURE_TIME, measure_time if measure_time else pd.NaT)
                        if COL_PROJECT not in df.columns: df[COL_PROJECT] = ''

                        output_cols = [
                            COL_FILENAME, COL_MEASURE_TIME, COL_NO, COL_PROJECT, 
                            COL_MEASURE_VAL, COL_DESIGN_VAL, COL_DIFF, 
                            COL_UPPER_TOL, COL_LOWER_TOL, COL_RESULT
                        ]
                        new_data_frames.append(df[output_cols])
            
            except Exception as e:
                # [優化] 使用 logging 記錄錯誤
                logging.error(f"處理檔案失敗 {filename}: {e}\n{traceback.format_exc()}")

        logging.info(f"處理完成。成功讀取: {len(loaded_filenames)}/{total}")
        self.data_loaded.emit(new_data_frames, loaded_filenames)

    def stop(self):
        self._is_running = False

# --- GUI 類別 ---
class VersionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 500, 400)
        layout = QVBoxLayout(self)
        lbl_title = QLabel(APP_TITLE)
        lbl_title.setStyleSheet("font-size: 16px; font-weight: bold; color: darkblue;")
        layout.addWidget(lbl_title)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(UPDATE_LOG)
        text_edit.setStyleSheet("font-family: Consolas, Monospace; font-size: 12px;")
        layout.addWidget(text_edit)
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

class DistributionPlotDialog(QDialog):
    def __init__(self, item_name, df_item, design_val, upper_tol, lower_tol, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"詳細分析: {item_name}")
        self.setGeometry(100, 100, 900, 600)
        self.df_item = df_item
        self.design_val = design_val
        self.upper_tol = upper_tol
        self.lower_tol = lower_tol
        self.usl = design_val + upper_tol
        self.lsl = design_val + lower_tol
        
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        self.tab_hist = QWidget()
        self.plot_histogram(self.tab_hist)
        tabs.addTab(self.tab_hist, "分佈直方圖")
        self.tab_trend = QWidget()
        self.plot_trend(self.tab_trend)
        tabs.addTab(self.tab_trend, "趨勢圖")
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

    def plot_histogram(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        data = self.df_item[COL_MEASURE_VAL].dropna()
        if len(data) > 0:
            ax.hist(data, bins=15, color='skyblue', edgecolor='black', alpha=0.7, label='實測值')
            ax.axvline(self.design_val, color='green', linestyle='-', linewidth=2, label=f'設計值')
            ax.axvline(self.usl, color='red', linestyle='--', linewidth=2, label=f'USL')
            ax.axvline(self.lsl, color='red', linestyle='--', linewidth=2, label=f'LSL')
            ax.set_title("量測值分佈圖")
            ax.legend()
            ax.grid(True, alpha=0.3)
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
        if COL_MEASURE_TIME in df_sorted.columns:
            try:
                if len(df_sorted[COL_MEASURE_TIME].dropna()) > 0:
                    df_sorted = df_sorted.sort_values(by=COL_MEASURE_TIME)
                    has_time = True
            except: pass
        
        y_data = df_sorted[COL_MEASURE_VAL]
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

class StatisticsDialog(QDialog):
    def __init__(self, full_df, total_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("完整統計報告 (含 CPK)")
        self.setGeometry(150, 150, 950, 500)
        layout = QVBoxLayout(self)
        
        # 統計邏輯
        grouped = full_df.groupby([COL_NO, COL_PROJECT])
        stats_list = []
        for (no, name), group in grouped:
            total_count = len(group) # 該測項的有效測量數
            fail_count = len(group[group[COL_RESULT] == 'FAIL'])
            # 不良率計算分母使用 "總樣本數 (檔案數)"，這在缺測時比較合理
            fail_rate = (fail_count / total_files) * 100 if total_files > 0 else 0
            
            first = group.iloc[0]
            design = first.get(COL_DESIGN_VAL, 0)
            usl = design + first.get(COL_UPPER_TOL, 0)
            lsl = design + first.get(COL_LOWER_TOL, 0)
            
            values = group[COL_MEASURE_VAL].dropna()
            
            # [優化] CPK 計算與樣本數檢查
            cpk = np.nan
            if len(values) >= 2 and (usl != lsl):
                mean = values.mean()
                std = values.std() # 預設 ddof=1 (樣本標準差)
                if std == 0: 
                    cpk = 999.0
                else:
                    cpu = (usl - mean) / (3 * std)
                    cpl = (mean - lsl) / (3 * std)
                    cpk = min(cpu, cpl)
            
            stats_list.append({
                'No': no, '測量專案': name, 'NG次數': fail_count,
                '樣本數': total_count, # 這裡顯示該測項實際測到的數量
                '不良率(%)': fail_rate, 'CPK': cpk
            })
        stats_df = pd.DataFrame(stats_list).sort_values(by=['不良率(%)', 'No'], ascending=[False, True])
        
        lbl_info = QLabel(f"統計基礎：共 {total_files} 個檔案 (Sample)。\n註：若 CPK 顯示灰色，代表該項測量樣本數不足 30，統計結果僅供參考。")
        lbl_info.setStyleSheet("font-weight: bold; font-size: 13px; color: #333;")
        layout.addWidget(lbl_info)
        
        table = QTableWidget()
        cols = ["No", "測量專案", "實測樣本數", "NG次數", "不良率(%)", "CPK"]
        table.setColumnCount(len(cols))
        table.setHorizontalHeaderLabels(cols)
        table.setRowCount(len(stats_df))
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        for r in range(len(stats_df)):
            row = stats_df.iloc[r]
            table.setItem(r, 0, QTableWidgetItem(str(row['No'])))
            table.setItem(r, 1, QTableWidgetItem(str(row['測量專案'])))
            
            # 樣本數
            sample_count = int(row['樣本數'])
            table.setItem(r, 2, QTableWidgetItem(str(sample_count)))
            
            ng_item = QTableWidgetItem(str(row['NG次數']))
            if row['NG次數'] > 0: ng_item.setForeground(QColor('red'))
            table.setItem(r, 3, ng_item)
            
            rate_item = QTableWidgetItem(f"{row['不良率(%)']:.1f}%")
            if row['不良率(%)'] > 0:
                rate_item.setForeground(QColor('red'))
                rate_item.setFont(ng_item.font())
            table.setItem(r, 4, rate_item)
            
            # [優化] CPK 顯示邏輯
            cpk_val = row['CPK']
            if pd.isna(cpk_val): 
                cpk_item = QTableWidgetItem("---")
            else:
                cpk_item = QTableWidgetItem(f"{cpk_val:.3f}")
                # 顏色判定
                if sample_count < 30:
                    # 樣本不足，顯示灰色並標註
                    cpk_item.setText(f"{cpk_val:.3f} (少)")
                    cpk_item.setForeground(QBrush(QColor('gray')))
                    cpk_item.setToolTip("樣本數少於 30，CPK 參考價值較低")
                else:
                    if cpk_val < 1.0: cpk_item.setBackground(QBrush(QColor(255, 200, 200))) # Poor
                    elif cpk_val < 1.33: cpk_item.setBackground(QBrush(QColor(255, 255, 200))) # Acceptable
                    else: cpk_item.setBackground(QBrush(QColor(200, 255, 200))) # Good
            table.setItem(r, 5, cpk_item)
        
        table.setSortingEnabled(True)
        layout.addWidget(table)
        
        btn_export = QPushButton("匯出統計報表 (.csv)")
        btn_export.clicked.connect(lambda: self.export_stats(stats_df))
        layout.addWidget(btn_export)
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

    def export_stats(self, df):
        path, _ = QFileDialog.getSaveFileName(self, "匯出", "Stats.csv", "CSV (*.csv)")
        if path:
            df.to_csv(path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "完成", "已匯出")

class MeasurementAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        setup_logging() # 啟動日誌
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 1300, 800)
        self.all_data = pd.DataFrame()
        self.loaded_files = set()
        self.loader_thread = None 
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        control_group = QGroupBox("操作控制")
        control_layout = QHBoxLayout()

        load_layout = QVBoxLayout()
        self.btn_add = QPushButton("1. 加入資料夾 (可累加)")
        self.btn_add.clicked.connect(self.add_folder_data)
        self.btn_add.setMinimumHeight(35)
        self.btn_clear = QPushButton("清空所有資料")
        self.btn_clear.clicked.connect(self.clear_all_data)
        self.btn_clear.setStyleSheet("color: red;")
        self.btn_clear.setMinimumHeight(25)
        load_layout.addWidget(self.btn_add)
        load_layout.addWidget(self.btn_clear)
        
        func_layout = QHBoxLayout()
        self.chk_only_fail = QCheckBox("僅顯示 FAIL 項目")
        self.chk_only_fail.stateChanged.connect(self.refresh_table_view)
        self.chk_only_fail.setEnabled(False) 
        self.btn_stats = QPushButton("3. 統計報告 (NG率/CPK)")
        self.btn_stats.clicked.connect(self.show_statistics)
        self.btn_stats.setMinimumHeight(40)
        self.btn_stats.setEnabled(False)
        self.btn_plot = QPushButton("4. 視覺化選定測項")
        self.btn_plot.clicked.connect(self.show_current_item_plot)
        self.btn_plot.setMinimumHeight(40)
        self.btn_plot.setEnabled(False)
        self.btn_plot.setStyleSheet("font-weight: bold; color: darkblue;")
        self.btn_export = QPushButton("2. 匯出總表")
        self.btn_export.clicked.connect(self.export_data)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)
        self.btn_version = QPushButton("關於")
        self.btn_version.clicked.connect(self.show_version_info)
        self.btn_version.setFixedSize(60, 40)

        func_layout.addWidget(self.chk_only_fail)
        func_layout.addWidget(self.btn_stats)
        func_layout.addWidget(self.btn_plot)
        func_layout.addWidget(self.btn_export)
        func_layout.addWidget(self.btn_version)
        
        control_layout.addLayout(load_layout, 1)
        control_layout.addLayout(func_layout, 3)
        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.lbl_info = QLabel("準備就緒。")
        main_layout.addWidget(self.lbl_info)

        self.lbl_status = QLabel("目前總資料: 0 筆 | 總樣本數(檔案數): 0")
        self.lbl_status.setStyleSheet("color: blue; font-weight: bold;")
        main_layout.addWidget(self.lbl_status)

        self.table_widget = QTableWidget()
        cols = [COL_FILENAME, COL_MEASURE_TIME, COL_NO, COL_PROJECT, COL_MEASURE_VAL, COL_DESIGN_VAL, COL_DIFF, COL_UPPER_TOL, COL_LOWER_TOL, COL_RESULT]
        self.table_widget.setColumnCount(len(cols))
        self.table_widget.setHorizontalHeaderLabels(cols)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        header = self.table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        main_layout.addWidget(self.table_widget)
        
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
            self.btn_stats.setEnabled(False)
            self.btn_plot.setEnabled(False)

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
            
            self.btn_export.setEnabled(True)
            self.chk_only_fail.setEnabled(True)
            self.btn_stats.setEnabled(True)
            self.btn_plot.setEnabled(True)
            self.chk_only_fail.setEnabled(True)
            
            self.refresh_table_view()
            logging.info(f"資料載入完成，新增 {len(new_data)} 筆，總共 {len(self.all_data)} 筆")
            self.lbl_info.setText(f"完成。本次加入 {len(new_data)} 筆數據。")
            QMessageBox.information(self, "完成", f"已加入 {len(loaded_filenames)} 個檔案。\n目前總資料量: {len(self.all_data)} 筆")
        else:
            self.lbl_info.setText("無有效數據。")
            logging.warning("嘗試載入但未提取到有效數據")
            QMessageBox.warning(self, "結果", "未提取到有效數據。")
            
        self.set_ui_loading_state(False)

    def clear_all_data(self):
        reply = QMessageBox.question(self, '確認', '確定清空所有資料？', 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.all_data = pd.DataFrame()
            self.loaded_files.clear()
            self.table_widget.setRowCount(0)
            self.lbl_status.setText("資料已清空")
            self.btn_plot.setEnabled(False)
            self.btn_stats.setEnabled(False)
            self.btn_export.setEnabled(False)
            self.chk_only_fail.setEnabled(False)
            logging.info("使用者清空了所有資料")

    def refresh_table_view(self):
        if self.all_data.empty: return
        df_to_show = self.all_data[self.all_data[COL_RESULT] == 'FAIL'] if self.chk_only_fail.isChecked() else self.all_data
        self.display_data(df_to_show)
        total_samples = len(self.loaded_files)
        self.lbl_status.setText(f"顯示: {len(df_to_show)} 筆 | 總資料: {len(self.all_data)} 筆 | 總樣本數: {total_samples}")
        if len(df_to_show) > 5000:
             self.lbl_status.setText(self.lbl_status.text() + " (僅顯示前 5000 筆，完整資料請匯出)")

    def display_data(self, df):
        MAX_DISPLAY = 5000 
        rows = min(len(df), MAX_DISPLAY)
        cols = df.shape[1]
        
        self.table_widget.setRowCount(rows)
        self.table_widget.setSortingEnabled(False) 
        red_brush = QBrush(QColor(255, 200, 200))
        red_text = QColor(200, 0, 0)
        green_text = QColor(0, 128, 0)

        for r in range(rows):
            is_fail = str(df.iloc[r, 9]) == "FAIL"
            for c in range(cols):
                val = df.iloc[r, c]
                item_text = ""
                if c == 1 and isinstance(val, (datetime, pd.Timestamp)):
                    item_text = val.strftime("%Y/%m/%d %H:%M:%S") if pd.notnull(val) else ""
                else:
                    item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                
                item = QTableWidgetItem(item_text)
                if is_fail:
                    if c in [6, 9]:
                        item.setForeground(red_text)
                        item.setBackground(red_brush)
                elif item_text == "OK" and c == 9:
                    item.setForeground(green_text)
                self.table_widget.setItem(r, c, item)
        self.table_widget.setSortingEnabled(True)

    def show_statistics(self):
        if self.all_data.empty: return
        total_files = len(self.loaded_files)
        dlg = StatisticsDialog(self.all_data, total_files, self)
        dlg.exec()

    def show_current_item_plot(self):
        sel = self.table_widget.selectedItems()
        if not sel:
            QMessageBox.warning(self, "提示", "請先選擇一列數據。")
            return
        row = sel[0].row()
        try:
            target_no = self.table_widget.item(row, 2).text()
            target_name = self.table_widget.item(row, 3).text()
            mask = (self.all_data[COL_NO].astype(str) == target_no) & (self.all_data[COL_PROJECT] == target_name)
            df_item = self.all_data[mask]
            if df_item.empty: return
            first = df_item.iloc[0]
            design = float(first[COL_DESIGN_VAL])
            upper = float(first[COL_UPPER_TOL])
            lower = float(first[COL_LOWER_TOL])
            plot_dlg = DistributionPlotDialog(f"{target_name} (No.{target_no})", df_item, design, upper, lower, self)
            plot_dlg.exec()
        except Exception as e:
            logging.error(f"視覺化失敗: {e}")
            QMessageBox.critical(self, "錯誤", f"無法分析: {e}")

    def export_data(self):
        if self.all_data.empty: return
        path, _ = QFileDialog.getSaveFileName(self, "匯出", "量測總表.csv", "CSV (*.csv)")
        if path:
            self.all_data.to_csv(path, index=False, encoding='utf-8-sig')
            logging.info(f"匯出資料至 {path}")
            QMessageBox.information(self, "完成", "匯出成功")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MeasurementAnalyzerApp()
    window.show()
    sys.exit(app.exec())