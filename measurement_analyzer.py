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
                             QTabWidget)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush, QAction

# Matplotlib imports for plotting
import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

# 設定中文字型 (嘗試尋找系統可用中文字型，避免亂碼)
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

class DistributionPlotDialog(QDialog):
    """
    視覺化分析視窗：顯示直方圖與趨勢圖
    """
    def __init__(self, item_name, df_item, design_val, upper_tol, lower_tol, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"詳細分析: {item_name}")
        self.setGeometry(100, 100, 900, 600)
        
        self.df_item = df_item
        self.design_val = design_val
        self.upper_tol = upper_tol
        self.lower_tol = lower_tol
        
        # 計算規格界限
        self.usl = design_val + upper_tol
        self.lsl = design_val + lower_tol
        
        layout = QVBoxLayout(self)
        
        # 建立分頁
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # 分頁 1: 直方圖
        self.tab_hist = QWidget()
        self.plot_histogram(self.tab_hist)
        tabs.addTab(self.tab_hist, "分佈直方圖 (Histogram)")
        
        # 分頁 2: 趨勢圖
        self.tab_trend = QWidget()
        self.plot_trend(self.tab_trend)
        tabs.addTab(self.tab_trend, "趨勢圖 (Trend)")
        
        # 關閉按鈕
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

    def plot_histogram(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        
        # 建立 Matplotlib Figure
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        
        ax = fig.add_subplot(111)
        
        # 繪製直方圖
        data = self.df_item['實測值'].dropna()
        if len(data) > 0:
            n, bins, patches = ax.hist(data, bins=15, color='skyblue', edgecolor='black', alpha=0.7, label='實測值')
            
            # 畫規格線
            ax.axvline(self.design_val, color='green', linestyle='-', linewidth=2, label=f'設計值 ({self.design_val})')
            ax.axvline(self.usl, color='red', linestyle='--', linewidth=2, label=f'USL ({self.usl})')
            ax.axvline(self.lsl, color='red', linestyle='--', linewidth=2, label=f'LSL ({self.lsl})')
            
            ax.set_title("量測值分佈圖")
            ax.set_xlabel("數值 (mm)")
            ax.set_ylabel("頻率 (次數)")
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
        
        # 嘗試依照時間排序 (如果有解析出時間)
        df_sorted = self.df_item.copy()
        if '測量時間' in df_sorted.columns and not df_sorted['測量時間'].isnull().all():
            try:
                df_sorted = df_sorted.sort_values(by='測量時間')
                x_label = "時間順序"
            except:
                x_label = "檔案讀取順序"
        else:
            x_label = "樣本索引"
            
        y_data = df_sorted['實測值']
        x_data = range(1, len(y_data) + 1)
        
        ax.plot(x_data, y_data, marker='o', linestyle='-', color='blue', markersize=4, label='實測值')
        
        # 畫規格線
        ax.axhline(self.design_val, color='green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        # 標示出超出規格的點
        fails = df_sorted[ (df_sorted['實測值'] > self.usl) | (df_sorted['實測值'] < self.lsl) ]
        if not fails.empty:
            # 找出這些點在原序列中的位置
            # 這裡簡單處理：重新對應 index
            # 實際應用建議用 scatter 覆蓋
            pass 

        ax.set_title("量測值趨勢圖")
        ax.set_xlabel(x_label)
        ax.set_ylabel("數值")
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        layout.addWidget(toolbar)
        layout.addWidget(canvas)


class StatisticsDialog(QDialog):
    """
    顯示統計結果的彈出視窗
    包含 Fail 率計算與 CPK
    """
    def __init__(self, full_df, total_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("完整統計報告 (含 CPK)")
        self.setGeometry(150, 150, 900, 500)
        
        layout = QVBoxLayout(self)
        
        # 1. 數據準備
        # 依 No, 測量專案 分組
        grouped = full_df.groupby(['No', '測量專案'])
        
        stats_list = []
        for (no, name), group in grouped:
            total_count = len(group)
            fail_count = len(group[group['判定結果'] == 'FAIL'])
            fail_rate = (fail_count / total_files) * 100 if total_files > 0 else 0
            
            # 取得設計值與公差 (取該組第一筆非空資料)
            first_valid = group.iloc[0]
            design = first_valid.get('設計值', 0)
            usl = design + first_valid.get('上限公差', 0)
            lsl = design + first_valid.get('下限公差', 0)
            
            # 計算 CPK
            values = group['實測值'].dropna()
            if len(values) > 1 and (usl != lsl):
                mean = values.mean()
                std = values.std()
                
                if std == 0:
                    cpk = 999.0 # 無變異，視為極好
                else:
                    cpu = (usl - mean) / (3 * std)
                    cpl = (mean - lsl) / (3 * std)
                    cpk = min(cpu, cpl)
            else:
                cpk = np.nan

            stats_list.append({
                'No': no,
                '測量專案': name,
                'NG次數': fail_count,
                '樣本數': total_files, # 總檔案數
                '不良率(%)': fail_rate,
                'CPK': cpk
            })
            
        stats_df = pd.DataFrame(stats_list)
        # 排序: 先排不良率(高->低)，再排 No
        stats_df = stats_df.sort_values(by=['不良率(%)', 'No'], ascending=[False, True])
        
        # 2. 介面顯示
        lbl_info = QLabel(f"統計基礎：共 {total_files} 個 CSV 檔案 (晶片)")
        lbl_info.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(lbl_info)
        
        table = QTableWidget()
        columns = ["No", "測量專案", "樣本數", "NG次數", "不良率(%)", "CPK"]
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(len(stats_df))
        
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        
        for r in range(len(stats_df)):
            row = stats_df.iloc[r]
            
            table.setItem(r, 0, QTableWidgetItem(str(row['No'])))
            table.setItem(r, 1, QTableWidgetItem(str(row['測量專案'])))
            table.setItem(r, 2, QTableWidgetItem(str(row['樣本數'])))
            
            # NG 次數
            ng_item = QTableWidgetItem(str(row['NG次數']))
            if row['NG次數'] > 0:
                ng_item.setForeground(QColor('red'))
            table.setItem(r, 3, ng_item)
            
            # 不良率
            rate_item = QTableWidgetItem(f"{row['不良率(%)']:.1f}%")
            if row['不良率(%)'] > 0:
                rate_item.setForeground(QColor('red'))
                rate_item.setFont(ng_item.font()) # 繼承字型
            table.setItem(r, 4, rate_item)
            
            # CPK
            cpk_val = row['CPK']
            if pd.isna(cpk_val):
                cpk_str = "---"
                bg_color = None
            else:
                cpk_str = f"{cpk_val:.3f}"
                # CPK 顏色判定
                if cpk_val < 1.0:
                    bg_color = QColor(255, 200, 200) # 紅 (Poor)
                elif cpk_val < 1.33:
                    bg_color = QColor(255, 255, 200) # 黃 (Acceptable)
                else:
                    bg_color = QColor(200, 255, 200) # 綠 (Good)
            
            cpk_item = QTableWidgetItem(cpk_str)
            if bg_color:
                cpk_item.setBackground(QBrush(bg_color))
            table.setItem(r, 5, cpk_item)

        table.setSortingEnabled(True)
        layout.addWidget(table)
        
        # 匯出統計按鈕
        btn_export = QPushButton("匯出統計報表 (.csv)")
        btn_export.clicked.connect(lambda: self.export_stats(stats_df))
        layout.addWidget(btn_export)
        
        btn_close = QPushButton("關閉")
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

    def export_stats(self, df):
        file_path, _ = QFileDialog.getSaveFileName(self, "匯出統計", "Fail統計與CPK.csv", "CSV Files (*.csv)")
        if file_path:
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "成功", "統計報表已匯出")


class MeasurementAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("量測數據分析工具 (Pro版: 統計/視覺化/CPK)")
        self.setGeometry(100, 100, 1300, 800)

        # 儲存數據
        self.all_data = pd.DataFrame()
        # 追蹤讀取過的檔案，計算總樣本數
        self.loaded_files = set()

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- 上方控制區 ---
        control_group = QGroupBox("操作控制")
        control_layout = QHBoxLayout()

        # 左側：載入
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
        
        # 右側：分析功能
        func_layout = QHBoxLayout()
        
        self.chk_only_fail = QCheckBox("僅顯示 FAIL 項目")
        self.chk_only_fail.stateChanged.connect(self.refresh_table_view)
        self.chk_only_fail.setEnabled(False) 

        self.btn_stats = QPushButton("3. 統計報告 (NG率/CPK)")
        self.btn_stats.clicked.connect(self.show_statistics)
        self.btn_stats.setMinimumHeight(40)
        self.btn_stats.setEnabled(False)

        # [新增] 視覺化按鈕
        self.btn_plot = QPushButton("4. 視覺化選定測項")
        self.btn_plot.setToolTip("請先在下方表格選擇某一列，再點擊此按鈕")
        self.btn_plot.clicked.connect(self.show_current_item_plot)
        self.btn_plot.setMinimumHeight(40)
        self.btn_plot.setEnabled(False)
        self.btn_plot.setStyleSheet("font-weight: bold; color: darkblue;")

        self.btn_export = QPushButton("2. 匯出總表")
        self.btn_export.clicked.connect(self.export_data)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)

        func_layout.addWidget(self.chk_only_fail)
        func_layout.addWidget(self.btn_stats)
        func_layout.addWidget(self.btn_plot) # Add to layout
        func_layout.addWidget(self.btn_export)
        
        control_layout.addLayout(load_layout, 1)
        control_layout.addLayout(func_layout, 3)
        control_group.setLayout(control_layout)
        main_layout.addWidget(control_group)

        # --- 進度與狀態 ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        self.lbl_info = QLabel("準備就緒。")
        main_layout.addWidget(self.lbl_info)

        self.lbl_status = QLabel("目前總資料: 0 筆 | 總樣本數(檔案數): 0")
        self.lbl_status.setStyleSheet("color: blue; font-weight: bold;")
        main_layout.addWidget(self.lbl_status)

        # --- 表格 ---
        self.table_widget = QTableWidget()
        columns = [
            "檔案名稱", "測量時間", "No", "測量專案", 
            "實測值", "設計值", "差異", 
            "上限公差", "下限公差", "判定結果"
        ]
        self.table_widget.setColumnCount(len(columns))
        self.table_widget.setHorizontalHeaderLabels(columns)
        self.table_widget.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows) # 整行選取
        self.table_widget.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        
        header = self.table_widget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        main_layout.addWidget(self.table_widget)

    def find_header_row_and_date(self, filepath):
        """尋找標題列以及嘗試解析日期"""
        header_idx = None
        encoding_found = None
        measure_time = None
        
        try:
            encodings = ['utf-8-sig', 'big5', 'cp950', 'shift_jis']
            for enc in encodings:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        lines = f.readlines()
                    
                    # 1. 找日期 (通常在前幾行)
                    # 格式範例: 測量日期及時間,2025/12/2 下午 03:08:11
                    for line in lines[:20]:
                        if "測量日期及時間" in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                time_str = parts[1].strip()
                                # 嘗試解析日期，這裡做簡單處理，如果不標準則保留字串
                                try:
                                    # 處理中文 '下午'/'上午' -> PM/AM 的轉換邏輯比較複雜，這裡保留原字串或做簡單替換
                                    # 若要排序精準，建議轉 datetime 物件
                                    measure_time = time_str
                                except:
                                    pass
                            break
                    
                    # 2. 找 Header
                    for i, line in enumerate(lines[:60]): 
                        if "No" in line and "實測值" in line and "設計值" in line:
                            header_idx = i
                            encoding_found = enc
                            return header_idx, encoding_found, measure_time
                    
                    # 如果找不到 header 但編碼沒報錯，繼續換編碼試試 (不太可能發生)
                    if header_idx is None: 
                        continue
                        
                except UnicodeDecodeError:
                    continue 
            return None, None, None
        except Exception as e:
            print(f"Error checking file: {e}")
            return None, None, None

    def add_folder_data(self):
        folder_path = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if not folder_path: return

        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        if not csv_files:
            QMessageBox.warning(self, "提示", "無 CSV 檔案。")
            return

        self.lbl_info.setText(f"處理中: {os.path.basename(folder_path)}...")
        self.progress_bar.setMaximum(len(csv_files))
        self.progress_bar.setValue(0)
        
        new_data_frames = []
        processed_count = 0
        
        for filepath in csv_files:
            try:
                # 取得 header 位置與日期
                header_idx, encoding, measure_time = self.find_header_row_and_date(filepath)
                
                if header_idx is not None:
                    # 記錄檔案名稱以計算總樣本數
                    filename = os.path.basename(filepath)
                    self.loaded_files.add(filename)
                    
                    df = pd.read_csv(
                        filepath, 
                        skiprows=header_idx, 
                        header=0, 
                        encoding=encoding, 
                        on_bad_lines='skip',
                        index_col=False 
                    )
                    
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # 欄位修正
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
                        
                        # 加入 metadata
                        df.insert(0, '檔案名稱', filename)
                        df.insert(1, '測量時間', measure_time if measure_time else "")
                        
                        if '測量專案' not in df.columns: df['測量專案'] = ''

                        output_cols = [
                            '檔案名稱', '測量時間', 'No', '測量專案', 
                            '實測值', '設計值', '差異', 
                            '上限公差', '下限公差', '判定結果'
                        ]
                        new_data_frames.append(df[output_cols])
            
            except Exception as e:
                print(f"Error {filepath}: {e}")
            
            processed_count += 1
            self.progress_bar.setValue(processed_count)
            QApplication.processEvents()

        if new_data_frames:
            new_data = pd.concat(new_data_frames, ignore_index=True)
            if self.all_data.empty:
                self.all_data = new_data
            else:
                self.all_data = pd.concat([self.all_data, new_data], ignore_index=True)
            
            self.btn_export.setEnabled(True)
            self.chk_only_fail.setEnabled(True)
            self.btn_stats.setEnabled(True)
            self.btn_plot.setEnabled(True)
            
            self.refresh_table_view()
            self.lbl_info.setText(f"完成加入。")
            QMessageBox.information(self, "完成", f"已加入 {len(new_data)} 筆數據。")
        else:
            QMessageBox.warning(self, "結果", "無有效數據。")

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

    def refresh_table_view(self):
        if self.all_data.empty: return

        if self.chk_only_fail.isChecked():
            df_to_show = self.all_data[self.all_data['判定結果'] == 'FAIL']
        else:
            df_to_show = self.all_data

        self.display_data(df_to_show)
        
        total_samples = len(self.loaded_files)
        self.lbl_status.setText(f"顯示: {len(df_to_show)} 筆 | 總資料: {len(self.all_data)} 筆 | 總樣本數(csv數): {total_samples}")

    def display_data(self, df):
        self.table_widget.setRowCount(0)
        if df.empty: return

        rows, cols = df.shape
        self.table_widget.setRowCount(rows)
        self.table_widget.setSortingEnabled(False) 

        red_brush = QBrush(QColor(255, 200, 200))
        red_text = QColor(200, 0, 0)
        green_text = QColor(0, 128, 0)

        # 這裡為了效能，可以只顯示前 2000 筆，或分頁 (如果數據量太大)
        # 目前先全顯
        for r in range(rows):
            is_fail = str(df.iloc[r, 9]) == "FAIL" # index 9 is '判定結果'

            for c in range(cols):
                val = df.iloc[r, c]
                item_text = f"{val:.4f}" if isinstance(val, float) else str(val)
                item = QTableWidgetItem(item_text)
                
                if is_fail:
                    if c in [6, 9]: # 差異, 判定
                        item.setForeground(red_text)
                        item.setBackground(red_brush)
                elif item_text == "OK" and c == 9:
                    item.setForeground(green_text)
                
                self.table_widget.setItem(r, c, item)

        self.table_widget.setSortingEnabled(True)

    def show_statistics(self):
        if self.all_data.empty: return
        # 傳入 總樣本數 (unique files)
        total_files = len(self.loaded_files)
        dlg = StatisticsDialog(self.all_data, total_files, self)
        dlg.exec()

    def show_current_item_plot(self):
        """顯示目前選定列的測項分佈圖"""
        selected_items = self.table_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "提示", "請先在表格中點選想要分析的那一列數據。")
            return
        
        # 取得選定列的 No 和 測量專案
        row = selected_items[0].row()
        
        # 表格可能有排序，所以要小心取得正確的數值
        # 這裡直接讀 UI 上的文字比較保險 (假設欄位順序沒變)
        # Columns: 0:檔名, 1:時間, 2:No, 3:專案, 4:實測, 5:設計, 6:差異, 7:上限, 8:下限
        try:
            target_no = self.table_widget.item(row, 2).text() # No
            target_name = self.table_widget.item(row, 3).text() # 專案
            
            # 從原始數據中篩選出所有符合此 No 和 Name 的資料
            # 必須轉換型別以匹配，DataFrame 中的 No 可能是 int 或 str
            # 這裡簡單用 string 比較
            mask = (self.all_data['No'].astype(str) == target_no) & (self.all_data['測量專案'] == target_name)
            df_item = self.all_data[mask]
            
            if df_item.empty:
                QMessageBox.warning(self, "錯誤", "找不到相關數據。")
                return

            # 取得規格 (假設同一測項規格都一樣，取第一筆)
            first = df_item.iloc[0]
            design = float(first['設計值'])
            upper = float(first['上限公差'])
            lower = float(first['下限公差'])
            
            # 開啟繪圖視窗
            plot_dlg = DistributionPlotDialog(f"{target_name} (No.{target_no})", df_item, design, upper, lower, self)
            plot_dlg.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"無法分析選定項目: {e}")

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