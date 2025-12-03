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
from PyQt6.QtGui import QColor, QBrush

# Matplotlib imports for plotting
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

# --- 設定常數 ---
APP_VERSION = "v1.5.0"
APP_TITLE = f"量測數據分析工具 (Pro版) {APP_VERSION}"
LOG_FILENAME = "measurement_analyzer.log"

UPDATE_LOG = """
=== 版本更新紀錄 ===

[v1.5.0] - 2025/12/03 (Stable Release)
1. [整合] 彙整所有先前測試功能，發布為穩定版 v1.5.0。
2. [系統] 完整的日誌 (Log) 追蹤系統，將錯誤寫入 log 檔以便排查。
3. [效能] 背景多執行緒 (QThread) 讀取 CSV，解決視窗凍結問題。
4. [功能] 支援 Keyence 中文日期格式解析 (上午/下午) 用於趨勢圖排序。
5. [統計] 完整的 CPK 計算 (含樣本數不足提示) 與 Fail 率統計。
6. [視覺] 整合直方圖與趨勢圖分析功能。

[v1.0.0 - v1.4.0] 歷史開發歷程
- 基礎 CSV 讀取與差異計算
- 公差判定邏輯 (略過設計值為 0 的項目)
- 視覺化圖表導入
- 多資料夾累加功能
"""

# --- 初始化日誌系統 ---
def setup_logging():
    """設定日誌系統，將輸出同時導向檔案與主控台"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILENAME, encoding='utf-8', mode='w'), # mode='w' 每次重啟清空舊log
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info(f"應用程式啟動 - {APP_TITLE}")

# --- 設定中文字型 (Matplotlib) ---
import matplotlib.font_manager as fm
def set_chinese_font():
    # 常見的 Windows/Mac/Linux 中文字型清單
    font_names = ['Microsoft JhengHei', 'SimHei', 'PingFang TC', 'Arial Unicode MS']
    for name in font_names:
        if name in [f.name for f in fm.fontManager.ttflist]:
            matplotlib.rcParams['font.sans-serif'] = [name]
            matplotlib.rcParams['axes.unicode_minus'] = False # 解決負號顯示問題
            break
set_chinese_font()

# --- 日期解析工具 ---
def parse_keyence_date(date_str):
    """
    解析 Keyence 特殊日期格式，例如: "2025/12/2 下午 03:08:11"
    """
    if not isinstance(date_str, str): return None
    date_str = date_str.strip()
    try:
        # 格式: YYYY/M/D [上午/下午] H:M:S
        match = re.match(r'(\d+)/(\d+)/(\d+)\s+(上午|下午)\s+(\d+):(\d+):(\d+)', date_str)
        if match:
            year, month, day, ampm, hour, minute, second = match.groups()
            year, month, day = int(year), int(month), int(day)
            hour, minute, second = int(hour), int(minute), int(second)
            
            # 12小時制轉24小時制
            if ampm == "下午" and hour < 12:
                hour += 12
            elif ampm == "上午" and hour == 12:
                hour = 0
                
            return datetime(year, month, day, hour, minute, second)
        else:
            # 嘗試解析標準格式
            formats = ["%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %p %I:%M:%S"]
            for fmt in formats:
                try:
                    return datetime.strptime(date_str, fmt)
                except ValueError:
                    continue
            return None
    except Exception as e:
        logging.warning(f"日期解析失敗 '{date_str}': {e}")
        return None

# --- 背景工作執行緒 (防止 UI 凍結) ---
class FileLoaderThread(QThread):
    progress_updated = pyqtSignal(int, str) # 更新進度條
    data_loaded = pyqtSignal(list, set)     # 回傳處理好的 DataFrame 列表
    error_occurred = pyqtSignal(str)        # 回傳嚴重錯誤

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
                        lines = f.readlines()
                    
                    measure_time = None
                    # 1. 找日期 (通常在前20行)
                    for line in lines[:20]:
                        if "測量日期及時間" in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                measure_time = parse_keyence_date(parts[1].strip())
                            break
                    
                    # 2. 找 Header
                    for i, line in enumerate(lines[:60]): 
                        if "No" in line and "實測值" in line and "設計值" in line:
                            return i, enc, measure_time
                except UnicodeDecodeError:
                    continue 
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
                    
                    # 清理欄位空白
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # 防呆: 修正 No 欄位
                    if 'No' not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: 'No'}, inplace=True)
                                break
                    
                    required_cols = ['No', '實測值', '設計值']
                    if all(col in df.columns for col in required_cols):
                        df = df.dropna(subset=['No'])
                        
                        # 轉數值
                        cols_to_numeric = ['實測值', '設計值', '上限公差', '下限公差']
                        for col in cols_to_numeric:
                            if col in df.columns:
                                df[col] = pd.to_numeric(df[col], errors='coerce')
                            else:
                                df[col] = 0.0
                        
                        df['差異'] = df['實測值'] - df['設計值']
                        
                        # 判定邏輯
                        def check_status(row):
                            if abs(row['設計值']) < 0.000001: return "---"
                            if pd.isna(row['上限公差']) or pd.isna(row['下限公差']): return "---"
                            # 如果公差都是0，且判斷欄位有值，則信任原檔判斷
                            if row['上限公差'] == 0 and row['下限公差'] == 0:
                                if '判斷' in row and pd.notna(row['判斷']): return row['判斷']
                                return "---"

                            diff = row['差異']
                            if diff > row['上限公差']: return "FAIL"
                            elif diff < row['下限公差']: return "FAIL"
                            else: return "OK"

                        df['判定結果'] = df.apply(check_status, axis=1)
                        
                        # 加入 Metadata
                        df.insert(0, '檔案名稱', filename)
                        df.insert(1, '測量時間', measure_time if measure_time else pd.NaT)
                        if '測量專案' not in df.columns: df['測量專案'] = ''

                        output_cols = [
                            '檔案名稱', '測量時間', 'No', '測量專案', 
                            '實測值', '設計值', '差異', 
                            '上限公差', '下限公差', '判定結果'
                        ]
                        new_data_frames.append(df[output_cols])
            
            except Exception as e:
                logging.error(f"檔案處理失敗 {filename}: {e}\n{traceback.format_exc()}")

        logging.info(f"Thread 處理完成。成功: {len(loaded_filenames)}/{total}")
        self.data_loaded.emit(new_data_frames, loaded_filenames)

    def stop(self):
        self._is_running = False

# --- 對話視窗類別 ---
class VersionDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 600, 450)
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
        
        data = self.df_item['實測值'].dropna()
        if len(data) > 0:
            ax.hist(data, bins=15, color='skyblue', edgecolor='black', alpha=0.7, label='實測值')
            ax.axvline(self.design_val, color='green', linestyle='-', linewidth=2, label=f'設計值 ({self.design_val})')
            ax.axvline(self.usl, color='red', linestyle='--', linewidth=2, label=f'USL')
            ax.axvline(self.lsl, color='red', linestyle='--', linewidth=2, label=f'LSL')
            ax.set_title("量測值分佈圖")
            ax.set_xlabel("數值 (mm)")
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
        
        # 嘗試依照時間排序
        if '測量時間' in df_sorted.columns:
            try:
                valid_times = df_sorted['測量時間'].dropna()
                if len(valid_times) > 0:
                    df_sorted = df_sorted.sort_values(by='測量時間')
                    has_time = True
            except: pass
        
        y_data = df_sorted['實測值']
        x_data = range(1, len(y_data) + 1)
        
        ax.plot(x_data, y_data, marker='o', linestyle='-', color='blue', markersize=4, label='實測值')
        ax.axhline(self.design_val, color='green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        ax.set_title("量測值趨勢圖")
        ax.set_xlabel("時間順序" if has_time else "讀取順序")
        ax.set_ylabel("數值")
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
        
        # 統計運算
        grouped = full_df.groupby(['No', '測量專案'])
        stats_list = []
        for (no, name), group in grouped:
            total_count = len(group)
            fail_count = len(group[group['判定結果'] == 'FAIL'])
            # 不良率分母為總檔案數(Sample數)，除非該測項在某些檔案中缺失
            fail_rate = (fail_count / total_files) * 100 if total_files > 0 else 0
            
            first = group.iloc[0]
            design = first.get('設計值', 0)
            usl = design + first.get('上限公差', 0)
            lsl = design + first.get('下限公差', 0)
            
            values = group['實測值'].dropna()
            
            # CPK 計算
            cpk = np.nan
            if len(values) >= 2 and (usl != lsl):
                mean = values.mean()
                std = values.std()
                if std == 0: 
                    cpk = 999.0
                else:
                    cpu = (usl - mean) / (3 * std)
                    cpl = (mean - lsl) / (3 * std)
                    cpk = min(cpu, cpl)
            
            stats_list.append({
                'No': no, '測量專案': name, 'NG次數': fail_count,
                '樣本數': total_count, # 實際有測到的數量
                '不良率(%)': fail_rate, 'CPK': cpk
            })
            
        stats_df = pd.DataFrame(stats_list).sort_values(by=['不良率(%)', 'No'], ascending=[False, True])
        
        lbl_info = QLabel(f"統計基礎：共 {total_files} 個樣本 (檔案)。")
        lbl_info.setStyleSheet("font-weight: bold; font-size: 13px; color: #333;")
        layout.addWidget(lbl_info)
        
        table = QTableWidget()
        cols = ["No", "測量專案", "實測數", "NG次數", "不良率(%)", "CPK"]
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
            
            # 實測數
            count_item = QTableWidgetItem(str(row['樣本數']))
            table.setItem(r, 2, count_item)
            
            # NG
            ng_item = QTableWidgetItem(str(row['NG次數']))
            if row['NG次數'] > 0: ng_item.setForeground(QColor('red'))
            table.setItem(r, 3, ng_item)
            
            # Rate
            rate_item = QTableWidgetItem(f"{row['不良率(%)']:.1f}%")
            if row['不良率(%)'] > 0:
                rate_item.setForeground(QColor('red'))
                rate_item.setFont(ng_item.font())
            table.setItem(r, 4, rate_item)
            
            # CPK (含顏色與樣本過少提示)
            cpk_val = row['CPK']
            if pd.isna(cpk_val): 
                cpk_item = QTableWidgetItem("---")
            else:
                cpk_item = QTableWidgetItem(f"{cpk_val:.3f}")
                if int(row['樣本數']) < 30:
                    cpk_item.setText(f"{cpk_val:.3f} (少)")
                    cpk_item.setForeground(QBrush(QColor('gray')))
                    cpk_item.setToolTip("樣本數 < 30，統計參考價值較低")
                else:
                    if cpk_val < 1.0: cpk_item.setBackground(QBrush(QColor(255, 200, 200))) # Red
                    elif cpk_val < 1.33: cpk_item.setBackground(QBrush(QColor(255, 255, 200))) # Yellow
                    else: cpk_item.setBackground(QBrush(QColor(200, 255, 200))) # Green
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

# --- 主應用程式 ---
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
        logging.info("主視窗初始化完成")

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
        cols = ["檔案名稱", "測量時間", "No", "測量專案", "實測值", "設計值", "差異", "上限公差", "下限公差", "判定結果"]
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

        # 啟動 Thread
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
            logging.info(f"資料載入完成，新增 {len(new_data)} 筆，目前總計 {len(self.all_data)} 筆")
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
        df_to_show = self.all_data[self.all_data['判定結果'] == 'FAIL'] if self.chk_only_fail.isChecked() else self.all_data
        self.display_data(df_to_show)
        
        total_samples = len(self.loaded_files)
        status_text = f"顯示: {len(df_to_show)} 筆 | 總資料: {len(self.all_data)} 筆 | 總樣本數: {total_samples}"
        
        # 分頁顯示提示
        if len(df_to_show) > 5000:
             status_text += " (僅顯示前 5000 筆，完整資料請匯出)"
        self.lbl_status.setText(status_text)

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
            is_fail = str(df.iloc[r, 9]) == "FAIL" # index 9 is '判定結果'
            for c in range(cols):
                val = df.iloc[r, c]
                item_text = ""
                # 日期格式化
                if c == 1 and isinstance(val, (datetime, pd.Timestamp)):
                    item_text = val.strftime("%Y/%m/%d %H:%M:%S") if pd.notnull(val) else ""
                else:
                    item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                
                item = QTableWidgetItem(item_text)
                if is_fail:
                    if c in [6, 9]: # 差異 & 判定欄位標紅
                        item.setForeground(red_text)
                        item.setBackground(red_brush)
                elif item_text == "OK" and c == 9:
                    item.setForeground(green_text)
                self.table_widget.setItem(r, c, item)
        self.table_widget.setSortingEnabled(True)

    def show_statistics(self):
        if self.all_data.empty: return
        total_files = len(self.loaded_files)
        try:
            dlg = StatisticsDialog(self.all_data, total_files, self)
            dlg.exec()
        except Exception as e:
            logging.error(f"統計報表錯誤: {e}")
            QMessageBox.critical(self, "錯誤", f"無法產生報表: {e}")

    def show_current_item_plot(self):
        sel = self.table_widget.selectedItems()
        if not sel:
            QMessageBox.warning(self, "提示", "請先選擇一列數據。")
            return
        row = sel[0].row()
        try:
            target_no = self.table_widget.item(row, 2).text()
            target_name = self.table_widget.item(row, 3).text()
            mask = (self.all_data['No'].astype(str) == target_no) & (self.all_data['測量專案'] == target_name)
            df_item = self.all_data[mask]
            
            if df_item.empty: return
            first = df_item.iloc[0]
            design = float(first['設計值'])
            upper = float(first['上限公差'])
            lower = float(first['下限公差'])
            
            plot_dlg = DistributionPlotDialog(f"{target_name} (No.{target_no})", df_item, design, upper, lower, self)
            plot_dlg.exec()
        except Exception as e:
            logging.error(f"視覺化失敗: {e}")
            QMessageBox.critical(self, "錯誤", f"無法分析: {e}")

    def export_data(self):
        if self.all_data.empty: return
        path, _ = QFileDialog.getSaveFileName(self, "匯出", "量測總表.csv", "CSV (*.csv)")
        if path:
            try:
                self.all_data.to_csv(path, index=False, encoding='utf-8-sig')
                logging.info(f"匯出資料至 {path}")
                QMessageBox.information(self, "完成", "匯出成功")
            except Exception as e:
                QMessageBox.critical(self, "匯出錯誤", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MeasurementAnalyzerApp()
    window.show()
    sys.exit(app.exec())