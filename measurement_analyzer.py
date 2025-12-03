import sys
import os
import glob
import pandas as pd
import numpy as np
import re
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
APP_VERSION = "v1.2.0"
APP_TITLE = f"量測數據分析工具 (Pro版) {APP_VERSION}"
UPDATE_LOG = """
=== 版本更新紀錄 ===

[v1.2.0] - 2025/12/03 (New)
1. [重構] 導入 QThread 多執行緒架構：
   - 解決讀取大量 CSV 檔案時視窗凍結的問題。
   - 優化記憶體管理，大幅提升檔案合併速度。
2. [優化] 進度條顯示更平滑，並新增讀取狀態提示。

[v1.1.0] - 2025/12/03
1. [修正] 日期解析邏輯：新增對 Keyence 報告中文日期格式 (上午/下午) 的支援。
2. [新增] 版本追蹤功能。
3. [優化] CSV 檔頭解析容錯率。

[v1.0.0] - 2025/12/01
1. 初始版本發布。
"""

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

# --- [新增] 背景工作執行緒 ---
class FileLoaderThread(QThread):
    # 定義訊號
    progress_updated = pyqtSignal(int, str) # 進度(int), 狀態文字(str)
    data_loaded = pyqtSignal(list, set)     # 完成後的資料列表(list), 成功讀取的檔名集合(set)
    error_occurred = pyqtSignal(str)        # 錯誤訊息

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._is_running = True

    def find_header_row_and_date(self, filepath):
        """將解析邏輯封裝在 Thread 內部"""
        try:
            encodings = ['utf-8-sig', 'big5', 'cp950', 'shift_jis']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        lines = f.readlines()
                    
                    measure_time = None
                    # 1. 找日期
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
        except:
            return None, None, None

    def run(self):
        new_data_frames = []
        loaded_filenames = set()
        total = len(self.file_paths)
        
        for i, filepath in enumerate(self.file_paths):
            if not self._is_running: break # 允許中斷

            filename = os.path.basename(filepath)
            # 發送進度訊號
            self.progress_updated.emit(i + 1, f"正在處理: {filename}")
            
            try:
                header_idx, encoding, measure_time = self.find_header_row_and_date(filepath)
                
                if header_idx is not None:
                    loaded_filenames.add(filename)
                    df = pd.read_csv(filepath, skiprows=header_idx, header=0, 
                                   encoding=encoding, on_bad_lines='skip', index_col=False)
                    
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # 欄位正規化
                    if 'No' not in df.columns:
                        for col in df.columns:
                            if 'No' in col and len(col) < 10:
                                df.rename(columns={col: 'No'}, inplace=True)
                                break
                    
                    required_cols = ['No', '實測值', '設計值']
                    if all(col in df.columns for col in required_cols):
                        df = df.dropna(subset=['No'])
                        
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
                            if row['上限公差'] == 0 and row['下限公差'] == 0:
                                if '判斷' in row and pd.notna(row['判斷']): return row['判斷']
                                return "---"
                            diff = row['差異']
                            if diff > row['上限公差']: return "FAIL"
                            elif diff < row['下限公差']: return "FAIL"
                            else: return "OK"

                        df['判定結果'] = df.apply(check_status, axis=1)
                        
                        # Metadata
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
                # 這裡不中斷，只記錄錯誤或忽略
                print(f"Error processing {filename}: {e}")

        # 完成後發送資料
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
        data = self.df_item['實測值'].dropna()
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
        if '測量時間' in df_sorted.columns:
            try:
                if len(df_sorted['測量時間'].dropna()) > 0:
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
        ax.legend()
        ax.grid(True, alpha=0.3)
        layout.addWidget(toolbar)
        layout.addWidget(canvas)

class StatisticsDialog(QDialog):
    def __init__(self, full_df, total_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("完整統計報告 (含 CPK)")
        self.setGeometry(150, 150, 900, 500)
        layout = QVBoxLayout(self)
        grouped = full_df.groupby(['No', '測量專案'])
        stats_list = []
        for (no, name), group in grouped:
            total_count = len(group)
            fail_count = len(group[group['判定結果'] == 'FAIL'])
            fail_rate = (fail_count / total_files) * 100 if total_files > 0 else 0
            first = group.iloc[0]
            design = first.get('設計值', 0)
            usl = design + first.get('上限公差', 0)
            lsl = design + first.get('下限公差', 0)
            values = group['實測值'].dropna()
            if len(values) > 1 and (usl != lsl):
                mean = values.mean()
                std = values.std()
                if std == 0: cpk = 999.0
                else:
                    cpu = (usl - mean) / (3 * std)
                    cpl = (mean - lsl) / (3 * std)
                    cpk = min(cpu, cpl)
            else: cpk = np.nan
            stats_list.append({
                'No': no, '測量專案': name, 'NG次數': fail_count,
                '樣本數': total_files, '不良率(%)': fail_rate, 'CPK': cpk
            })
        stats_df = pd.DataFrame(stats_list).sort_values(by=['不良率(%)', 'No'], ascending=[False, True])
        
        lbl_info = QLabel(f"統計基礎：共 {total_files} 個 CSV 檔案 (晶片)")
        lbl_info.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(lbl_info)
        table = QTableWidget()
        cols = ["No", "測量專案", "樣本數", "NG次數", "不良率(%)", "CPK"]
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
            table.setItem(r, 2, QTableWidgetItem(str(row['樣本數'])))
            ng_item = QTableWidgetItem(str(row['NG次數']))
            if row['NG次數'] > 0: ng_item.setForeground(QColor('red'))
            table.setItem(r, 3, ng_item)
            rate_item = QTableWidgetItem(f"{row['不良率(%)']:.1f}%")
            if row['不良率(%)'] > 0:
                rate_item.setForeground(QColor('red'))
                rate_item.setFont(ng_item.font())
            table.setItem(r, 4, rate_item)
            cpk_val = row['CPK']
            if pd.isna(cpk_val): 
                cpk_item = QTableWidgetItem("---")
            else:
                cpk_item = QTableWidgetItem(f"{cpk_val:.3f}")
                if cpk_val < 1.0: cpk_item.setBackground(QBrush(QColor(255, 200, 200)))
                elif cpk_val < 1.33: cpk_item.setBackground(QBrush(QColor(255, 255, 200)))
                else: cpk_item.setBackground(QBrush(QColor(200, 255, 200)))
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
        self.setWindowTitle(APP_TITLE)
        self.setGeometry(100, 100, 1300, 800)
        self.all_data = pd.DataFrame()
        self.loaded_files = set()
        self.loader_thread = None # 用於存放 Thread 實例
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

        # 鎖定按鈕避免重複操作
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
        """控制按鈕狀態"""
        self.btn_add.setEnabled(not is_loading)
        self.btn_clear.setEnabled(not is_loading)
        if is_loading:
            self.btn_export.setEnabled(False)
            self.btn_stats.setEnabled(False)
            self.btn_plot.setEnabled(False)

    def on_progress_updated(self, value, message):
        """接收 Thread 的進度更新"""
        self.progress_bar.setValue(value)
        self.lbl_info.setText(message)

    def on_data_loaded(self, new_data_frames, loaded_filenames):
        """接收 Thread 處理完成的資料"""
        self.loaded_files.update(loaded_filenames)
        
        if new_data_frames:
            self.lbl_info.setText("正在合併資料...")
            QApplication.processEvents() # 讓介面更新一下
            
            # 一次性合併 (效能遠優於迴圈合併)
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
            self.lbl_info.setText(f"完成。本次加入 {len(new_data)} 筆數據。")
            QMessageBox.information(self, "完成", f"已加入 {len(loaded_filenames)} 個檔案。\n目前總資料量: {len(self.all_data)} 筆")
        else:
            self.lbl_info.setText("無有效數據。")
            QMessageBox.warning(self, "結果", "未提取到有效數據。")
            
        # 解鎖 UI
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

    def refresh_table_view(self):
        if self.all_data.empty: return
        df_to_show = self.all_data[self.all_data['判定結果'] == 'FAIL'] if self.chk_only_fail.isChecked() else self.all_data
        self.display_data(df_to_show)
        total_samples = len(self.loaded_files)
        self.lbl_status.setText(f"顯示: {len(df_to_show)} 筆 | 總資料: {len(self.all_data)} 筆 | 總樣本數: {total_samples}")

    def display_data(self, df):
        # 限制顯示數量以避免 UI 卡頓 (若資料過大)
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
            QMessageBox.critical(self, "錯誤", f"無法分析: {e}")

    def export_data(self):
        if self.all_data.empty: return
        path, _ = QFileDialog.getSaveFileName(self, "匯出", "量測總表.csv", "CSV (*.csv)")
        if path:
            self.all_data.to_csv(path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "完成", "匯出成功")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MeasurementAnalyzerApp()
    window.show()
    sys.exit(app.exec())