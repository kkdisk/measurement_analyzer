# -*- coding: utf-8 -*-
"""
Measurement Analyzer - 主程式入口
整合所有模組並啟動 GUI 應用程式
"""
import sys
import os
import glob
import logging
import pandas as pd
import numpy as np
from datetime import datetime

# PyQt6 imports
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHeaderView, QProgressBar, QMessageBox, QGroupBox, QCheckBox, 
                             QInputDialog, QAbstractItemView, QTabWidget, QHBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

# Internal imports
from config import AppConfig, DISPLAY_COLUMNS
from statistics import calculate_cpk, calculate_tolerance_for_yield
from parsers import natural_keys, HAS_PDF_SUPPORT
from widgets import NumericTableWidgetItem, VersionDialog, DistributionPlotDialog
from workers import FileLoaderThread

# Optional Theme Support
try:
    import qdarktheme
    HAS_THEME_SUPPORT = True
except ImportError:
    HAS_THEME_SUPPORT = False

# Natsort
try:
    from natsort import index_natsorted, ns
    HAS_NATSORT = True
except ImportError:
    HAS_NATSORT = False


def setup_logging():
    """初始化日誌系統"""
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
                    f"【達成 {AppConfig.DEFAULT_TARGET_YIELD*100:.0f}% 良率所需公差】",
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
