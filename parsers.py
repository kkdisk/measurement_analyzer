# -*- coding: utf-8 -*-
"""
Measurement Analyzer - 資料解析模組
包含 CSV/PDF 解析邏輯與日期處理
"""
import re
import pandas as pd
import logging
import traceback
import sys
from datetime import datetime
from config import AppConfig

# Optional PDF Support
try:
    import pdfplumber
    HAS_PDF_SUPPORT = True
except ImportError:
    HAS_PDF_SUPPORT = False


def natural_keys(text):
    """
    Fallback for natural sorting if natsort is missing.
    """
    try:
        text = str(text)
        return tuple([int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)])
    except Exception:
        return (str(text),)


def parse_keyence_date(date_str):
    """
    解析 Keyence 報告中的日期格式
    """
    if not isinstance(date_str, str): return None
    date_str = date_str.strip()
    try:
        # 處理 Keyence 常見格式 "2023/01/01 下午 01:23:45"
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
    except Exception: return None


def extract_text_by_clustering(page, y_tolerance=3):
    """
    使用座標聚類法提取 PDF 文字，解決表格錯位問題
    """
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


def read_pdf_file(filepath):
    """
    讀取 Keyence PDF 報告
    """
    if not HAS_PDF_SUPPORT: return None, None
    
    df = pd.DataFrame()
    measure_time = None
    data_list = []
    
    try:
        with pdfplumber.open(filepath) as pdf:
            first_page_words = extract_text_by_clustering(pdf.pages[0])
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
                lines = extract_text_by_clustering(page)
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


def find_header_row_and_date_csv(filepath):
    """
    尋找 CSV 檔頭與日期 (自動偵測編碼)
    """
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
            except Exception: continue 
        return None, None, None
    except Exception: return None, None, None
