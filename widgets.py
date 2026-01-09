# -*- coding: utf-8 -*-
"""
Measurement Analyzer - GUI 元件模組
包含自定義表格元件、對話框與圖表繪製
"""
import logging
import pandas as pd
import numpy as np

# PyQt6 imports
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                             QDialog, QTabWidget, QTextEdit, QTableWidgetItem, QGroupBox, QComboBox)
from PyQt6.QtGui import QColor, QBrush, QFont
from PyQt6.QtCore import Qt

# Matplotlib imports
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar

# Internal imports
from config import AppConfig, UPDATE_LOG
from parsers import natural_keys
from statistics import calculate_tolerance_for_yield

# Natsort
try:
    from natsort import natsort_keygen, ns
    HAS_NATSORT = True
except ImportError:
    HAS_NATSORT = False


def set_chinese_font():
    """設定 Matplotlib 中文字型 (回歸 v1.7.1 策略)"""
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


class NumericTableWidgetItem(QTableWidgetItem):
    """支援數值排序與自然排序的表格項目"""
    def __lt__(self, other):
        if HAS_NATSORT:
            try:
                natsort_key = natsort_keygen(alg=ns.IGNORECASE)
                return natsort_key(self.text()) < natsort_key(other.text())
            except Exception:
                pass
        
        try:
            # Fallback
            return natural_keys(self.text()) < natural_keys(other.text())
        except Exception:
            return super().__lt__(other)


class VersionDialog(QDialog):
    """版本資訊對話框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("版本資訊")
        self.setGeometry(300, 300, 600, 450)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(AppConfig.TITLE))
        txt = QTextEdit()
        txt.setReadOnly(True)
        txt.setPlainText(UPDATE_LOG)
        layout.addWidget(txt)
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)


class DistributionPlotDialog(QDialog):
    """詳細分佈與趨勢分析圖表對話框"""
    def __init__(self, item_name, df_item, design_val, upper_tol, lower_tol, parent=None, theme='light'):
        super().__init__(parent)
        self.setWindowTitle(f"詳細分析: {item_name}")
        self.setGeometry(100, 100, 950, 650)
        self.item_name = item_name
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
        
        # [v2.3.0] 新增公差建議分頁
        self.tab_tolerance = QWidget()
        self.setup_tolerance_tab(self.tab_tolerance)
        tabs.addTab(self.tab_tolerance, "📐 公差建議")
        
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)
    
    def setup_tolerance_tab(self, parent_widget):
        """設定公差建議分頁"""
        layout = QVBoxLayout(parent_widget)
        
        # 目標良率選擇區
        yield_group = QGroupBox("目標良率設定")
        yield_layout = QHBoxLayout()
        
        yield_layout.addWidget(QLabel("目標良率："))
        
        self.yield_combo = QComboBox()
        self.yield_combo.addItems(["80%", "85%", "90%", "95%", "99%", "99.73% (3σ)"])
        self.yield_combo.setCurrentIndex(2)  # 預設 90%
        self.yield_combo.currentIndexChanged.connect(self.update_tolerance_display)
        yield_layout.addWidget(self.yield_combo)
        
        yield_layout.addStretch()
        yield_group.setLayout(yield_layout)
        layout.addWidget(yield_group)
        
        # 結果顯示區
        result_group = QGroupBox("計算結果")
        result_layout = QVBoxLayout()
        
        self.tol_result_text = QTextEdit()
        self.tol_result_text.setReadOnly(True)
        self.tol_result_text.setMinimumHeight(300)
        result_layout.addWidget(self.tol_result_text)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 初始計算
        self.update_tolerance_display()
    
    def update_tolerance_display(self):
        """更新公差計算結果顯示"""
        yield_map = {0: 0.80, 1: 0.85, 2: 0.90, 3: 0.95, 4: 0.99, 5: 0.9973}
        target_yield = yield_map.get(self.yield_combo.currentIndex(), 0.90)
        
        vals = pd.to_numeric(self.df_item[AppConfig.Columns.MEASURED], errors='coerce').dropna()
        result = calculate_tolerance_for_yield(vals, self.design_val, target_yield)
        
        # 格式化輸出
        lines = []
        lines.append(f"═══════════════════════════════════════")
        lines.append(f"  測量專案：{self.item_name}")
        lines.append(f"  目標良率：{target_yield * 100:.2f}%")
        lines.append(f"═══════════════════════════════════════")
        lines.append("")
        
        if result['reliability'] == 'invalid':
            lines.append("❌ 無法計算：數據不足 (需至少 2 個樣本)")
        elif result['reliability'] == 'zero_std':
            lines.append("❌ 無法計算：標準差為零 (所有數據相同)")
        else:
            lines.append("📊 【數據統計】")
            lines.append(f"   樣本數：{len(vals)}")
            lines.append(f"   平均值 (μ)：{result['mean']:.4f}")
            lines.append(f"   標準差 (σ)：{result['std']:.4f}")
            lines.append(f"   設計值：{self.design_val:.4f}")
            lines.append(f"   製程偏移：{result['offset']:+.4f}")
            lines.append("")
            
            lines.append("📐 【建議公差】")
            lines.append(f"   ✅ 對稱公差：±{result['symmetric_tol']:.4f}")
            lines.append("")
            lines.append(f"   📈 非對稱建議：")
            lines.append(f"      上限公差：+{result['upper_tol']:.4f}")
            lines.append(f"      下限公差：{result['lower_tol']:.4f}")
            lines.append("")
            
            lines.append("📋 【與當前規格比較】")
            lines.append(f"   當前上限：+{self.upper_tol:.4f}")
            lines.append(f"   當前下限：{self.lower_tol:.4f}")
            
            current_max_tol = max(abs(self.upper_tol), abs(self.lower_tol))
            if current_max_tol > 0:
                ratio = result['symmetric_tol'] / current_max_tol
                if ratio > 1.2:
                    lines.append("")
                    lines.append(f"   ⚠️ 警告：要達到 {target_yield*100:.0f}% 良率，")
                    lines.append(f"      建議公差比當前規格大 {(ratio-1)*100:.1f}%")
                    lines.append(f"      建議放寬規格或改善製程")
                elif ratio < 0.8:
                    lines.append("")
                    lines.append(f"   ✅ 良好：當前規格充裕，")
                    lines.append(f"      實際只需 {ratio*100:.1f}% 即可達標")
                else:
                    lines.append("")
                    lines.append(f"   ℹ️ 規格適中 (比例：{ratio*100:.1f}%)")
            
            if result['reliability'] == 'small_sample':
                lines.append("")
                lines.append("⚠️ 注意：樣本數少於 30，結果僅供參考")
                lines.append("   建議累積更多數據後再做決策")
        
        self.tol_result_text.setPlainText("\n".join(lines))

    def plot_histogram(self, parent_widget):
        layout = QVBoxLayout(parent_widget)
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        data = self.df_item[AppConfig.Columns.MEASURED].dropna()
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
        if AppConfig.Columns.TIME in df_sorted.columns:
            try:
                if len(df_sorted[AppConfig.Columns.TIME].dropna()) > 0:
                    df_sorted = df_sorted.sort_values(by=AppConfig.Columns.TIME)
                    has_time = True
            except: pass
        
        y_data = df_sorted[AppConfig.Columns.MEASURED].values
        x_data = np.arange(1, len(y_data) + 1)
        
        # Prepare data for tooltip
        filenames = df_sorted[AppConfig.Columns.FILE].values if AppConfig.Columns.FILE in df_sorted.columns else []
        times = df_sorted[AppConfig.Columns.TIME].values if AppConfig.Columns.TIME in df_sorted.columns else []
        
        line_color = 'cyan' if self.theme == 'dark' else 'blue'
        line, = ax.plot(x_data, y_data, marker='o', linestyle='-', color=line_color, markersize=4, label='實測值')
        
        ax.axhline(self.design_val, color='lime' if self.theme=='dark' else 'green', linestyle='-', alpha=0.5, label='設計值')
        ax.axhline(self.usl, color='red', linestyle='--', alpha=0.5, label='USL')
        ax.axhline(self.lsl, color='red', linestyle='--', alpha=0.5, label='LSL')
        
        ax.set_title("量測值趨勢圖")
        ax.set_xlabel("時間順序" if has_time else "讀取順序")
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        # --- Tooltip Implementation ---
        annot = ax.annotate("", xy=(0,0), xytext=(10,10),textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="w", alpha=0.9),
                            arrowprops=dict(arrowstyle="->"))
        annot.set_visible(False)

        def update_annot(ind):
            x, y = line.get_data()
            idx = ind["ind"][0]
            annot.xy = (x[idx], y[idx])
            
            val = y_data[idx]
            fname = filenames[idx] if len(filenames) > idx else "Unknown"
            
            time_str = ""
            if len(times) > idx:
                t = times[idx]
                if pd.notnull(t):
                    try:
                        time_str = pd.to_datetime(t).strftime("%Y/%m/%d %H:%M:%S")
                    except: pass
            
            # Format text
            text = f"File: {fname}\nValue: {val:.4f}"
            if time_str:
                text += f"\nTime: {time_str}"
                
            annot.set_text(text)

        def hover(event):
            vis = annot.get_visible()
            if event.inaxes == ax:
                cont, ind = line.contains(event)
                if cont:
                    update_annot(ind)
                    annot.set_visible(True)
                    canvas.draw_idle()
                else:
                    if vis:
                        annot.set_visible(False)
                        canvas.draw_idle()

        canvas.mpl_connect("motion_notify_event", hover)
        # ------------------------------

        layout.addWidget(toolbar)
        layout.addWidget(canvas)
