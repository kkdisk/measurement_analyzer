# -*- coding: utf-8 -*-
"""
Measurement Analyzer - 背景工作模組
包含檔案載入執行緒
"""
import os
import pandas as pd
import logging
import traceback
from PyQt6.QtCore import QThread, pyqtSignal

from config import AppConfig, DISPLAY_COLUMNS
from parsers import find_header_row_and_date_csv, read_pdf_file

class FileLoaderThread(QThread):
    """檔案讀取背景執行緒"""
    progress_updated = pyqtSignal(int, str)
    data_loaded = pyqtSignal(list, set)
    error_occurred = pyqtSignal(str)

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._is_running = True

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
            except Exception:
                size_str = "Unknown"
                
            self.progress_updated.emit(i + 1, f"處理中: {filename} ({size_str})")
            
            df = None
            measure_time = None
            try:
                ext = os.path.splitext(filename)[1].lower()
                if ext == '.pdf':
                    df, measure_time = read_pdf_file(filepath)
                else:
                    header_idx, encoding, measure_time = find_header_row_and_date_csv(filepath)
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
