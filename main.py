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
from widgets import NumericTableWidgetItem, VersionDialog, DistributionPlotDialog, XYScatterPlotDialog, ArrayHeatmapDialog
from workers import FileLoaderThread
from xy_analyzer import classify_project_name, MeasurementType, get_xy_group_id

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
        
        # [v2.5.0] 控制列：合併 2D 顯示選項
        control_layout = QHBoxLayout()
        lbl_hint = QLabel("提示：雙擊表格任一列可開啟詳細圖表分析")
        lbl_hint.setStyleSheet("color: gray; font-style: italic;")
        control_layout.addWidget(lbl_hint)
        
        control_layout.addStretch()
        
        self.chk_merge_2d = QCheckBox("合併 2D XY 座標顯示")
        self.chk_merge_2d.setToolTip("勾選後，同一座標組的 X/Y 將合併為一行，顯示徑向偏差統計")
        self.chk_merge_2d.setChecked(True)  # [v2.5.0] 預設勾選，用戶可自行取消
        self.chk_merge_2d.stateChanged.connect(self.on_merge_2d_changed)
        control_layout.addWidget(self.chk_merge_2d)
        
        layout.addLayout(control_layout)

        self.stats_table = QTableWidget()
        cols = ["No", "測量專案", "類型", "樣本數", "NG數", "不良率(%)", "CPK", "建議公差(90%)", "平均值", "最大值", "最小值"]
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

    def on_merge_2d_changed(self, state):
        """[v2.5.0] 合併 2D 顯示切換事件處理"""
        # 重新計算並刷新統計表
        if not self.all_data.empty:
            self.calculate_and_refresh_stats()

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
        
        # [v2.5.0] 合併 2D XY 座標顯示邏輯
        merge_2d = self.chk_merge_2d.isChecked()
        processed_xy_groups = set()  # 已處理的 XY 座標組
        xy_group_data = {}  # 收集 XY 座標組資料用於合併統計
        
        stats_list = []
        for (no, name), group in grouped:
            # [v2.5.0] 分類測量類型
            type_info, group_id, axis = classify_project_name(name)
            type_label = type_info.value
            
            # 若勾選合併 2D，收集 XY 資料稍後處理
            if merge_2d and type_info == MeasurementType.XY_COORD:
                if group_id not in xy_group_data:
                    xy_group_data[group_id] = {'x_group': None, 'y_group': None, 'no': no}
                if axis == 'X':
                    xy_group_data[group_id]['x_group'] = group
                else:
                    xy_group_data[group_id]['y_group'] = group
                continue  # 跳過獨立 X/Y，稍後處理合併
            
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
                "No": no, "測量專案": name, "類型": type_label, "樣本數": count, 
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
        
        # [v2.5.0] 處理合併 XY 座標組統計
        if merge_2d and xy_group_data:
            from xy_analyzer import calculate_radial_deviation, calculate_radial_tolerance
            
            for group_id, data in xy_group_data.items():
                x_group = data.get('x_group')
                y_group = data.get('y_group')
                
                if x_group is None or y_group is None:
                    continue
                
                # 按檔案配對計算徑向偏差
                x_by_file = {row[AppConfig.Columns.FILE]: row for _, row in x_group.iterrows()}
                y_by_file = {row[AppConfig.Columns.FILE]: row for _, row in y_group.iterrows()}
                
                radial_devs = []
                ng_count = 0
                first_row = None
                
                for file_name, x_row in x_by_file.items():
                    if file_name not in y_by_file:
                        continue
                    y_row = y_by_file[file_name]
                    first_row = x_row  # 用於取得公差等資訊
                    
                    x_val = pd.to_numeric(x_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    x_design = pd.to_numeric(x_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    y_val = pd.to_numeric(y_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    y_design = pd.to_numeric(y_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    
                    if any(np.isnan([x_val, x_design, y_val, y_design])):
                        continue
                    
                    dx = x_val - x_design
                    dy = y_val - y_design
                    radial = calculate_radial_deviation(dx, dy)
                    radial_devs.append(radial)
                    
                    # 判定徑向是否超標
                    upper_tol = pd.to_numeric(x_row.get(AppConfig.Columns.UPPER), errors='coerce')
                    radial_tol = calculate_radial_tolerance(upper_tol, upper_tol)
                    if not np.isnan(radial_tol) and radial > radial_tol:
                        ng_count += 1
                
                if not radial_devs:
                    continue
                
                # 計算統計
                count = len(radial_devs)
                radial_arr = np.array(radial_devs)
                mean_radial = radial_arr.mean()
                max_radial = radial_arr.max()
                min_radial = radial_arr.min()
                std_radial = radial_arr.std(ddof=1) if count > 1 else 0
                fail_rate = (ng_count / total_files) * 100 if total_files > 0 else 0
                
                # [v2.5.0] 計算 2D CPK (CPU)
                cpu = np.nan
                if not np.isnan(radial_tol) and radial_tol > 0 and std_radial > 0:
                    cpu = (radial_tol - mean_radial) / (3 * std_radial)
                
                # [v2.5.0] 計算 2D 建議公差
                from xy_analyzer import calculate_2d_suggested_tolerance
                sugg_result = calculate_2d_suggested_tolerance(radial_arr, AppConfig.DEFAULT_TARGET_YIELD)
                sugg_tol = sugg_result.get('suggested_tol', np.nan)
                
                # 加入合併後的統計
                stats_list.append({
                    "No": data['no'],
                    "測量專案": f"{group_id} (2D合併)",
                    "類型": "2D",
                    "樣本數": count,
                    "NG數": ng_count,
                    "不良率(%)": fail_rate,
                    "CPK": cpu,  # 顯示 CPU
                    "CPK_RELIABILITY": 'ok' if count >= 30 else 'low_sample',
                    "建議公差": sugg_tol,
                    "TOL_RELIABILITY": sugg_result.get('reliability', 'invalid'),
                    "TOL_UPPER": radial_tol, # 顯示目前的徑向公差作為參考? 
                                             # 原本 TOL_UPPER 是用來顯示建議公差的上限.
                                             # 欄位定義: 建議公差 (數值).
                                             # 這裡填入 sugg_tol.
                    "TOL_LOWER": 0,
                    "TOL_OFFSET": 0,
                    "平均值": mean_radial,  # 顯示平均徑向偏差
                    "最大值": max_radial,
                    "最小值": min_radial,
                    "_design": 0,
                    "_upper": pd.to_numeric(first_row.get(AppConfig.Columns.UPPER), errors='coerce') if first_row is not None else 0,
                    "_lower": 0,
                    "_is_merged_2d": True  # 標記為合併項目，用於 Phase 3 展開
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
            
            # [v2.5.0] 類型欄 (1D/2D/陣列)
            type_item = QTableWidgetItem(str(row.get('類型', '1D')))
            if row.get('類型') == '2D':
                type_item.setForeground(QColor('blue'))
            elif row.get('類型') == '陣列':
                type_item.setForeground(QColor('purple'))
            self.stats_table.setItem(r, 2, type_item)
            
            self.stats_table.setItem(r, 3, NumericTableWidgetItem(str(row['樣本數'])))
            
            ng_item = NumericTableWidgetItem(str(row['NG數']))
            if row['NG數'] > 0: ng_item.setForeground(QColor('red'))
            self.stats_table.setItem(r, 4, ng_item)
            
            rate_item = NumericTableWidgetItem(f"{row['不良率(%)']:.2f}")
            if row['不良率(%)'] > 0: rate_item.setForeground(QColor('red'))
            self.stats_table.setItem(r, 5, rate_item)
            
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

            self.stats_table.setItem(r, 6, cpk_item)
            
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
            
            self.stats_table.setItem(r, 7, tol_item)
            self.stats_table.setItem(r, 8, NumericTableWidgetItem(f"{row['平均值']:.4f}"))
            self.stats_table.setItem(r, 9, NumericTableWidgetItem(f"{row['最大值']:.4f}"))
            self.stats_table.setItem(r, 10, NumericTableWidgetItem(f"{row['最小值']:.4f}"))
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
            # [v2.5.0] 檢查是否為陣列類型
            type_info, group_id, sub_info = classify_project_name(name)
            if type_info == MeasurementType.ARRAY:
                # 收集陣列所有點的資料
                unique_projects = self.all_data[AppConfig.Columns.PROJECT].unique()
                array_items = []
                
                for proj in unique_projects:
                    t, g, idx = classify_project_name(proj)
                    if t == MeasurementType.ARRAY and g == group_id:
                        # 計算該點的統計值 (顯示平均值)
                        # 限定目前的 No (雖然 No 通常不同檔案相同?) 
                        # 修正: No 應該是批號? 但主要以 Name 分組
                        # 這裡假設查看的是整個 DataSet 的平均表現
                        
                        mask = (self.all_data[AppConfig.Columns.PROJECT] == proj)
                        if no: # 如果有點選特定 No，是否要只過濾該 No? 
                               # 通常 No 是一批資料的 ID. 如果合併多個 CSV，No 可能不同?
                               # 原始邏輯 open_plot_dialog 傳入 no (如 '59', '60').
                               # 在 raw_table 中，No 是每一行的 ID. 
                               # 若要看整體分佈，應該忽略 No，或只看特定 No?
                               # 統計表是 aggregation. raw table 是 individual.
                               # open_plot_dialog 其實是用於查看 "Raw Data Row" 的詳情?
                               # 還是 "Stats Row"?
                               # 調用來源 plot_from_stats_table 傳入的是 row's No. (Group Key)
                               # 如果是 groupby(No, Project)，那麼只看該 No 的資料是正確的.
                            mask &= (self.all_data[AppConfig.Columns.NO].astype(str) == str(no))
                            
                        subset = self.all_data[mask]
                        vals = pd.to_numeric(subset[AppConfig.Columns.MEASURED], errors='coerce').dropna()
                        
                        if not vals.empty:
                            array_items.append({
                                'index': idx,
                                'value': vals.mean(), # 顯示平均值
                                'file': 'Average'
                            })
                
                # 排序 (數字優先)
                try:
                    array_items.sort(key=lambda x: int(x['index']) if isinstance(x['index'], int) or str(x['index']).isdigit() else str(x['index']))
                except:
                    array_items.sort(key=lambda x: str(x['index']))
                
                if array_items:
                    dlg = ArrayHeatmapDialog(group_id, array_items, self, self.current_theme)
                    dlg.exec()
                    return
            
            # [v2.5.0] 處理合併 2D 項目
            if " (2D合併)" in name:
                from xy_analyzer import calculate_radial_deviation, calculate_radial_tolerance
                
                group_id = name.replace(" (2D合併)", "")
                # 查找對應的 X/Y 原始資料
                x_name = f"{group_id}[X座標]"
                y_name = f"{group_id}[Y座標]"
                
                mask_x = (self.all_data[AppConfig.Columns.NO].astype(str) == no) & \
                         (self.all_data[AppConfig.Columns.PROJECT] == x_name)
                mask_y = (self.all_data[AppConfig.Columns.NO].astype(str) == no) & \
                         (self.all_data[AppConfig.Columns.PROJECT] == y_name)
                
                df_x = self.all_data[mask_x]
                df_y = self.all_data[mask_y]
                
                if df_x.empty or df_y.empty:
                    QMessageBox.information(self, "提示", f"找不到 {group_id} 的完整 X/Y 資料")
                    return
                
                # 按檔案配對計算徑向偏差
                x_by_file = {row[AppConfig.Columns.FILE]: row for _, row in df_x.iterrows()}
                y_by_file = {row[AppConfig.Columns.FILE]: row for _, row in df_y.iterrows()}
                
                radial_data = []
                first_x_row = None
                
                for file_name, x_row in x_by_file.items():
                    if file_name not in y_by_file:
                        continue
                    y_row = y_by_file[file_name]
                    first_x_row = x_row
                    
                    x_val = pd.to_numeric(x_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    x_design = pd.to_numeric(x_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    y_val = pd.to_numeric(y_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    y_design = pd.to_numeric(y_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    
                    if any(np.isnan([x_val, x_design, y_val, y_design])):
                        continue
                    
                    dx = x_val - x_design
                    dy = y_val - y_design
                    radial = calculate_radial_deviation(dx, dy)
                    
                    # 建立徑向偏差資料行
                    radial_data.append({
                        AppConfig.Columns.FILE: file_name,
                        AppConfig.Columns.NO: no,
                        AppConfig.Columns.PROJECT: f"{group_id} (徑向偏差)",
                        AppConfig.Columns.MEASURED: radial,  # 徑向偏差作為實測值
                        AppConfig.Columns.DESIGN: 0,  # 設計值為 0（期望中心點）
                        AppConfig.Columns.UPPER: x_row.get(AppConfig.Columns.UPPER, 0.05),  # 使用 X 的公差
                        AppConfig.Columns.LOWER: 0,  # 徑向偏差為正值
                        AppConfig.Columns.RESULT: 'OK'
                    })
                
                if not radial_data:
                    QMessageBox.information(self, "提示", "無法計算徑向偏差")
                    return
                
                # 取得公差資訊
                upper_tol = pd.to_numeric(first_x_row.get(AppConfig.Columns.UPPER, 0.05), errors='coerce')
                radial_tol = calculate_radial_tolerance(upper_tol, upper_tol)
                
                # [v2.5.0] 建立 XY 數據用於散佈圖
                xy_scatter_data = []
                for file_name, x_row in x_by_file.items():
                    if file_name not in y_by_file:
                        continue
                    y_row = y_by_file[file_name]
                    
                    x_val = pd.to_numeric(x_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    x_design = pd.to_numeric(x_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    y_val = pd.to_numeric(y_row.get(AppConfig.Columns.MEASURED), errors='coerce')
                    y_design = pd.to_numeric(y_row.get(AppConfig.Columns.DESIGN), errors='coerce')
                    
                    if any(np.isnan([x_val, x_design, y_val, y_design])):
                        continue
                    
                    dx = x_val - x_design
                    dy = y_val - y_design
                    radial = calculate_radial_deviation(dx, dy)
                    is_ng = radial > radial_tol if not np.isnan(radial_tol) else False
                    
                    xy_scatter_data.append({
                        'dx': dx,
                        'dy': dy,
                        'file': file_name,
                        'is_ng': is_ng
                    })
                
                # 開啟 2D 散佈圖對話框
                plot_dlg = XYScatterPlotDialog(
                    group_id,
                    xy_scatter_data,
                    radial_tol,
                    self,
                    self.current_theme
                )
                plot_dlg.exec()
            else:
                # 原有邏輯
                mask = (self.all_data[AppConfig.Columns.NO].astype(str) == no) & \
                       (self.all_data[AppConfig.Columns.PROJECT] == name)
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
