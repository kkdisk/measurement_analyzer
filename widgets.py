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
                             QDialog, QTabWidget, QTextEdit, QTableWidgetItem, QGroupBox, QComboBox, QDoubleSpinBox)
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
from xy_analyzer import calculate_2d_suggested_tolerance

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


class XYScatterPlotDialog(QDialog):
    """[v2.5.0] 2D XY 散佈圖對話框"""
    
    def __init__(self, group_name, xy_data, radial_tolerance, parent=None, theme='light'):
        """
        Args:
            group_name: 座標組名稱 (如 'NO.1_XY座標')
            xy_data: List of dicts with keys: 'dx', 'dy', 'file', 'is_ng'
            radial_tolerance: 徑向公差
            parent: 父視窗
            theme: 主題 ('light' or 'dark')
        """
        super().__init__(parent)
        self.setWindowTitle(f"2D 位置分佈圖: {group_name}")
        self.setGeometry(100, 100, 800, 700)
        self.group_name = group_name
        self.xy_data = xy_data
        self.radial_tolerance = radial_tolerance
        self.theme = theme
        
        # 設定 Style
        if self.theme == 'dark':
            plt.style.use('dark_background')
        else:
            plt.style.use('default')
        set_chinese_font()
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # 散佈圖頁籤
        self.tab_scatter = QWidget()
        self.plot_scatter(self.tab_scatter)
        tabs.addTab(self.tab_scatter, "📍 XY 分佈圖")
        
        # [v2.5.0] 徑向偏差直方圖頁籤
        self.tab_hist = QWidget()
        self.plot_radial_histogram(self.tab_hist)
        tabs.addTab(self.tab_hist, "📊 分佈直方圖")
        
        # [v2.5.0] 趨勢圖頁籤
        self.tab_trend = QWidget()
        self.plot_radial_trend(self.tab_trend)
        tabs.addTab(self.tab_trend, "📈 趨勢圖")
        
        # 統計摘要頁籤
        self.tab_stats = QWidget()
        self.setup_stats_tab(self.tab_stats)
        tabs.addTab(self.tab_stats, "📋 統計摘要")
        
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)
    
    def plot_scatter(self, parent_widget):
        """繪製 2D 散佈圖 (含手動公差設定)"""
        layout = QVBoxLayout(parent_widget)
        
        # [v2.5.0] 公差控制區
        ctrl_layout = QHBoxLayout()
        ctrl_layout.addWidget(QLabel("徑向公差 (mm):"))
        self.spin_tol = QDoubleSpinBox()
        self.spin_tol.setRange(0.0001, 100.0)
        self.spin_tol.setDecimals(4)
        self.spin_tol.setSingleStep(0.005)
        # 設定初始值 (若為 nan 則設為 0)
        init_tol = self.radial_tolerance if not np.isnan(self.radial_tolerance) else 0.05
        self.spin_tol.setValue(init_tol)
        
        btn_update = QPushButton("更新判定")
        btn_update.clicked.connect(self.update_scatter_plot)
        
        # [v2.5.0] 快速轉換按鈕
        btn_convert = QPushButton("轉為內切圓 (÷√2)")
        btn_convert.setToolTip("將公差除以 1.414 (從矩形對角轉為單軸標準)")
        btn_convert.clicked.connect(self.convert_tolerance_to_inscribed)
        
        ctrl_layout.addWidget(self.spin_tol)
        ctrl_layout.addWidget(btn_convert)
        ctrl_layout.addWidget(btn_update)
        ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)
        
        # 繪圖區
        self.fig_scatter = Figure(figsize=(7, 7), dpi=100)
        self.canvas_scatter = FigureCanvas(self.fig_scatter)
        self.toolbar_scatter = NavigationToolbar(self.canvas_scatter, parent_widget)
        
        layout.addWidget(self.toolbar_scatter)
        layout.addWidget(self.canvas_scatter)
        
        self.draw_scatter()

    def convert_tolerance_to_inscribed(self):
        """將目前公差值除以 sqrt(2)"""
        current_val = self.spin_tol.value()
        new_val = current_val / np.sqrt(2)
        self.spin_tol.setValue(new_val)
        self.update_scatter_plot()

    def update_scatter_plot(self):
        """更新公差並重繪"""
        new_tol = self.spin_tol.value()
        self.radial_tolerance = new_tol
        
        # 重新計算 NG 狀態
        # 假設 xy_data 中 dx, dy 單位已是 mm (或與公差一致)
        ng_count = 0
        for d in self.xy_data:
            r = np.sqrt(d['dx']**2 + d['dy']**2)
            is_ng = r > new_tol
            d['is_ng'] = is_ng
            if is_ng: ng_count += 1
            
        self.draw_scatter()
        
        # [Optional] 更新標題或其他資訊已反映新的 NG 數
        # self.setWindowTitle(f"2D 位置分佈圖: {self.group_name} (NG: {ng_count})")

    def draw_scatter(self):
        """執行繪圖邏輯"""
        self.fig_scatter.clear()
        ax = self.fig_scatter.add_subplot(111)
        
        # 設定等比例軸
        ax.set_aspect('equal', adjustable='box')
        
        # 繪製公差圓
        tol = self.radial_tolerance
        if tol > 0:
            # 公差圓（綠色填充）
            circle = plt.Circle((0, 0), tol, color='lightgreen', alpha=0.3, label=f'公差圓 (r={tol:.4f})')
            ax.add_patch(circle)
            # 公差圓邊界
            circle_edge = plt.Circle((0, 0), tol, color='green', fill=False, linewidth=2)
            ax.add_patch(circle_edge)
        
        # 繪製數據點
        dx_ok = [d['dx'] for d in self.xy_data if not d.get('is_ng', False)]
        dy_ok = [d['dy'] for d in self.xy_data if not d.get('is_ng', False)]
        dx_ng = [d['dx'] for d in self.xy_data if d.get('is_ng', False)]
        dy_ng = [d['dy'] for d in self.xy_data if d.get('is_ng', False)]
        
        # 計算新的比例
        total = len(self.xy_data)
        ok_ratio = len(dx_ok) / total * 100 if total > 0 else 0
        ng_ratio = len(dx_ng) / total * 100 if total > 0 else 0
        
        if dx_ok:
            ax.scatter(dx_ok, dy_ok, c='blue', s=50, alpha=0.7, label=f'合格: {len(dx_ok)} ({ok_ratio:.1f}%)', zorder=5)
        if dx_ng:
            ax.scatter(dx_ng, dy_ng, c='red', s=80, alpha=0.9, marker='x', label=f'超標: {len(dx_ng)} ({ng_ratio:.1f}%)', zorder=6)
        
        # 繪製原點標記
        ax.scatter([0], [0], c='green', s=100, marker='+', linewidths=2, label='設計中心', zorder=7)
        
        # 繪製座標軸
        ax.axhline(0, color='gray', linestyle='--', linewidth=0.5, alpha=0.5)
        ax.axvline(0, color='gray', linestyle='--', linewidth=0.5, alpha=0.5)
        
        # 設定範圍（確保能看到所有點和公差圓）
        all_dx = [d['dx'] for d in self.xy_data]
        all_dy = [d['dy'] for d in self.xy_data]
        if all_dx and all_dy:
            max_range = max(max(abs(min(all_dx)), abs(max(all_dx)), tol),
                           max(abs(min(all_dy)), abs(max(all_dy)), tol)) * 1.3
            ax.set_xlim(-max_range, max_range)
            ax.set_ylim(-max_range, max_range)
        
        ax.set_xlabel('X 偏差 (ΔX)')
        ax.set_ylabel('Y 偏差 (ΔY)')
        # self.group_name 可能包含 " (2D合併)"，視情況簡化
        ax.set_title(f'{self.group_name} - XY 位置分佈')
        ax.legend(loc='upper right')
        ax.grid(True, alpha=0.3)
        
        self.canvas_scatter.draw()
    
    def plot_radial_histogram(self, parent_widget):
        """[v2.5.0] 繪製徑向偏差直方圖"""
        layout = QVBoxLayout(parent_widget)
        
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        # 計算徑向偏差
        radial_vals = np.array([np.sqrt(d['dx']**2 + d['dy']**2) for d in self.xy_data])
        
        if len(radial_vals) > 0:
            color = 'cyan' if self.theme == 'dark' else 'skyblue'
            edgecolor = 'white' if self.theme == 'dark' else 'black'
            ax.hist(radial_vals, bins=15, color=color, edgecolor=edgecolor, alpha=0.7, label='徑向偏差')
            
            # 繪製公差線
            if self.radial_tolerance > 0:
                ax.axvline(self.radial_tolerance, color='red', linestyle='--', linewidth=2, 
                          label=f'徑向公差 ({self.radial_tolerance:.4f})')
            
            # 繪製平均線
            ax.axvline(radial_vals.mean(), color='lime' if self.theme=='dark' else 'green', 
                      linestyle='-', linewidth=2, label=f'平均 ({radial_vals.mean():.4f})')
            
            ax.set_title("徑向偏差分佈圖")
            ax.set_xlabel("徑向偏差")
            ax.set_ylabel("次數")
            ax.legend()
            ax.grid(True, alpha=0.3)
        else:
            ax.text(0.5, 0.5, "無有效數據", ha='center', va='center')
        
        layout.addWidget(toolbar)
        layout.addWidget(canvas)
    
    def plot_radial_trend(self, parent_widget):
        """[v2.5.0] 繪製徑向偏差趨勢圖"""
        layout = QVBoxLayout(parent_widget)
        
        fig = Figure(figsize=(8, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        # 計算徑向偏差
        radial_vals = np.array([np.sqrt(d['dx']**2 + d['dy']**2) for d in self.xy_data])
        filenames = [d.get('file', '') for d in self.xy_data]
        x_data = np.arange(1, len(radial_vals) + 1)
        
        if len(radial_vals) > 0:
            line_color = 'cyan' if self.theme == 'dark' else 'blue'
            line, = ax.plot(x_data, radial_vals, marker='o', linestyle='-', color=line_color, 
                           markersize=4, label='徑向偏差')
            
            # 繪製公差線
            if self.radial_tolerance > 0:
                ax.axhline(self.radial_tolerance, color='red', linestyle='--', alpha=0.5, 
                          label=f'徑向公差 ({self.radial_tolerance:.4f})')
            
            # 繪製平均線
            ax.axhline(radial_vals.mean(), color='lime' if self.theme=='dark' else 'green', 
                      linestyle='-', alpha=0.5, label=f'平均 ({radial_vals.mean():.4f})')
            
            ax.set_title("徑向偏差趨勢圖")
            ax.set_xlabel("樣本序號")
            ax.set_ylabel("徑向偏差")
            ax.legend()
            ax.grid(True, alpha=0.3)
            
            # Tooltip
            annot = ax.annotate("", xy=(0,0), xytext=(10,10), textcoords="offset points",
                               bbox=dict(boxstyle="round", fc="w", alpha=0.9),
                               arrowprops=dict(arrowstyle="->"))
            annot.set_visible(False)
            
            def update_annot(ind):
                idx = ind["ind"][0]
                annot.xy = (x_data[idx], radial_vals[idx])
                fname = filenames[idx] if idx < len(filenames) else "Unknown"
                text = f"File: {fname}\nRadial: {radial_vals[idx]:.4f}"
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
        else:
            ax.text(0.5, 0.5, "無有效數據", ha='center', va='center')
        
        layout.addWidget(toolbar)
        layout.addWidget(canvas)

    def setup_stats_tab(self, parent_widget):
        """設定統計摘要頁籤"""
        layout = QVBoxLayout(parent_widget)
        
        txt = QTextEdit()
        txt.setReadOnly(True)
        
        # 計算統計
        n = len(self.xy_data)
        dx_vals = np.array([d['dx'] for d in self.xy_data])
        dy_vals = np.array([d['dy'] for d in self.xy_data])
        radial_vals = np.sqrt(dx_vals**2 + dy_vals**2)
        ng_count = sum(1 for d in self.xy_data if d.get('is_ng', False))
        
        lines = []
        lines.append("═══════════════════════════════════════")
        lines.append(f"  2D 位置分析：{self.group_name}")
        lines.append("═══════════════════════════════════════")
        lines.append("")
        lines.append("📊 【樣本統計】")
        lines.append(f"   樣本數：{n}")
        lines.append(f"   NG 數：{ng_count}")
        lines.append(f"   不良率：{ng_count/n*100:.2f}%" if n > 0 else "   不良率：---")
        lines.append("")
        lines.append("📍 【X 軸偏差】")
        lines.append(f"   平均：{dx_vals.mean():.4f}")
        lines.append(f"   標準差：{dx_vals.std():.4f}")
        lines.append(f"   範圍：{dx_vals.min():.4f} ~ {dx_vals.max():.4f}")
        lines.append("")
        lines.append("📍 【Y 軸偏差】")
        lines.append(f"   平均：{dy_vals.mean():.4f}")
        lines.append(f"   標準差：{dy_vals.std():.4f}")
        lines.append(f"   範圍：{dy_vals.min():.4f} ~ {dy_vals.max():.4f}")
        lines.append("")
        lines.append("📐 【徑向偏差】")
        lines.append(f"   平均：{radial_vals.mean():.4f}")
        lines.append(f"   最大：{radial_vals.max():.4f}")
        lines.append(f"   最小：{radial_vals.min():.4f}")
        lines.append(f"   標準差：{radial_vals.std():.4f}")
        lines.append("")
        lines.append(f"   徑向公差：{self.radial_tolerance:.4f}")
        
        # 單側 CPK (CPU)
        if n > 1 and radial_vals.std() > 0:
            cpu = (self.radial_tolerance - radial_vals.mean()) / (3 * radial_vals.std())
            lines.append("")
            lines.append("📈 【2D CPK (CPU)】")
            lines.append(f"   CPU = (USL - μ) / (3σ)")
            lines.append(f"   CPU = ({self.radial_tolerance:.4f} - {radial_vals.mean():.4f}) / (3 × {radial_vals.std():.4f})")
            lines.append(f"   CPU = {cpu:.3f}")
            if cpu >= 1.33:
                lines.append("   ✅ 製程能力優良 (CPU ≥ 1.33)")
            elif cpu >= 1.0:
                lines.append("   ⚠️ 製程能力尚可 (1.0 ≤ CPU < 1.33)")
            else:
                lines.append("   ❌ 製程能力不足 (CPU < 1.0)")
        
        # [v2.5.0] 2D 建議公差
        sugg_result = calculate_2d_suggested_tolerance(radial_vals, target_yield=0.90)
        sugg_tol = sugg_result.get('suggested_tol', np.nan)
        
        lines.append("")
        lines.append("🛡️ 【建議公差】 (目標良率 90%, Rayleigh模型)")
        if not np.isnan(sugg_tol):
             lines.append(f"   建議徑向公差：{sugg_tol:.4f}")
             if self.radial_tolerance > 0:
                 ratio = sugg_tol / self.radial_tolerance
                 if ratio > 1.0:
                     lines.append(f"   ⚠️ 需放寬至當前規格的 {ratio*100:.1f}%")
                 else:
                     lines.append(f"   ✅ 當前規格充足 (只需 {ratio*100:.1f}%)")
        else:
             lines.append("   無法計算 (數據不足)")
        
        txt.setPlainText("\n".join(lines))
        layout.addWidget(txt)


class ArrayHeatmapDialog(QDialog):
    """[v2.5.0] 陣列資料視覺化對話框 (熱力圖/條形圖)"""
    
    def __init__(self, group_name, array_data, parent=None, theme='light'):
        """
        Args:
            group_name: 群組名稱 (如 'AA區平面度')
            array_data: List of dicts with keys: 'index', 'value', 'file'
            parent: 父視窗
            theme: 主題 ('light' or 'dark')
        """
        super().__init__(parent)
        self.setWindowTitle(f"陣列分析: {group_name}")
        self.setGeometry(100, 100, 900, 600)
        self.group_name = group_name
        self.array_data = array_data
        self.theme = theme
        
        # 設定 Style
        if self.theme == 'dark':
            plt.style.use('dark_background')
        else:
            plt.style.use('default')
        set_chinese_font()
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        tabs = QTabWidget()
        layout.addWidget(tabs)
        
        # 條形圖頁籤
        self.tab_bar = QWidget()
        self.plot_bar_chart(self.tab_bar)
        tabs.addTab(self.tab_bar, "📊 條形圖")
        
        # [v2.5.0] 熱力圖頁籤
        self.tab_heatmap = QWidget()
        self.plot_heatmap_ui(self.tab_heatmap)
        tabs.addTab(self.tab_heatmap, "🌡️ 2D 熱力圖")
        
        # 統計摘要頁籤
        self.tab_stats = QWidget()
        self.setup_stats_tab(self.tab_stats)
        tabs.addTab(self.tab_stats, "📋 統計摘要")
        
        btn = QPushButton("關閉")
        btn.clicked.connect(self.close)
        layout.addWidget(btn)
        
    def plot_bar_chart(self, parent_widget):
        """繪製數值條形圖"""
        layout = QVBoxLayout(parent_widget)
        
        fig = Figure(figsize=(10, 6), dpi=100)
        canvas = FigureCanvas(fig)
        toolbar = NavigationToolbar(canvas, parent_widget)
        ax = fig.add_subplot(111)
        
        # 準備數據
        indices = [d['index'] for d in self.array_data]
        values = [d['value'] for d in self.array_data]
        
        if len(values) > 0:
            # 顏色映射
            norm = plt.Normalize(min(values), max(values))
            cmap = plt.cm.get_cmap('coolwarm')
            colors = cmap(norm(values))
            
            bars = ax.bar(range(len(values)), values, color=colors, alpha=0.8)
            ax.set_xticks(range(len(values)))
            
            # 若點數太多，簡化 X 軸標籤
            if len(indices) > 30:
                n = len(indices)
                step = n // 20
                ax.set_xticks(range(0, n, step))
                ax.set_xticklabels([indices[i] for i in range(0, n, step)], rotation=45)
            else:
                ax.set_xticklabels(indices, rotation=45)
            
            # 標記 Max/Min
            min_idx = np.argmin(values)
            max_idx = np.argmax(values)
            
            ax.annotate(f'Min: {values[min_idx]:.3f}', 
                        xy=(min_idx, values[min_idx]), 
                        xytext=(0, -20), textcoords='offset points', ha='center',
                        arrowprops=dict(arrowstyle="->", color='blue'))
                        
            ax.annotate(f'Max: {values[max_idx]:.3f}', 
                        xy=(max_idx, values[max_idx]), 
                        xytext=(0, 20), textcoords='offset points', ha='center',
                        arrowprops=dict(arrowstyle="->", color='red'))
            
            ax.set_title(f"{self.group_name} - 各點平均值分佈")
            ax.set_ylabel("數值")
            ax.grid(True, alpha=0.3, axis='y')
            
            # Colorbar
            sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
            sm.set_array([])
            fig.colorbar(sm, ax=ax, label='數值')
            
        else:
            ax.text(0.5, 0.5, "無數據", ha='center', va='center')
            
        layout.addWidget(toolbar)
        layout.addWidget(canvas)

    def plot_heatmap_ui(self, parent_widget):
        """建立熱力圖 UI"""
        layout = QVBoxLayout(parent_widget)
        
        # 控制區
        ctrl_layout = QHBoxLayout()
        ctrl_layout.addWidget(QLabel("排列方式 (Rows x Cols):"))
        
        self.spin_rows = QComboBox()
        self.spin_cols = QComboBox()
        
        # 自動猜測維度
        N = len(self.array_data)
        factors = []
        for i in range(1, int(np.sqrt(N)) + 1):
            if N % i == 0:
                factors.append((i, N // i))
        
        # 預設邏輯 (優先 22x14, 14x22)
        default_idx = 0
        self.grid_options = []
        
        if N == 308:
            self.grid_options.append((22, 14))
            self.grid_options.append((14, 22))
        
        for r, c in factors:
            if (r,c) not in self.grid_options: self.grid_options.append((r, c))
            if (c,r) not in self.grid_options and r != c: self.grid_options.append((c, r))
            
        for r, c in self.grid_options:
            self.spin_rows.addItem(f"{r} x {c}")
            
        self.spin_rows.currentIndexChanged.connect(self.update_heatmap)
        
        ctrl_layout.addWidget(self.spin_rows)
        ctrl_layout.addStretch()
        layout.addLayout(ctrl_layout)
        
        # 繪圖區
        self.fig_hm = Figure(figsize=(8, 6), dpi=100)
        self.canvas_hm = FigureCanvas(self.fig_hm)
        self.toolbar_hm = NavigationToolbar(self.canvas_hm, parent_widget)
        
        layout.addWidget(self.toolbar_hm)
        layout.addWidget(self.canvas_hm)
        
        # 初始繪製
        self.update_heatmap()
        
    def update_heatmap(self):
        """更新熱力圖"""
        self.fig_hm.clear()
        ax = self.fig_hm.add_subplot(111)
        
        idx = self.spin_rows.currentIndex()
        if idx < 0 or idx >= len(self.grid_options):
            return
            
        rows, cols = self.grid_options[idx]
        
        values = [d['value'] for d in self.array_data]
        # 確保數據依照 index 排序 (由小到大)
        # 假設 array_data 已經排序過
        
        try:
            matrix = np.array(values).reshape(rows, cols)
            
            im = ax.imshow(matrix, cmap='coolwarm', interpolation='nearest') # 或 'bilinear'
            
            # Colorbar
            self.fig_hm.colorbar(im, ax=ax)
            
            # 添加數值標籤 (如果格子夠少)
            if len(values) < 100:
                for i in range(rows):
                    for j in range(cols):
                        val = matrix[i, j]
                        text = ax.text(j, i, f"{val:.1f}",
                                       ha="center", va="center", color="w", fontsize=8)
            
            ax.set_title(f"熱力圖 ({rows}x{cols}) - 所有樣本平均值")
            self.canvas_hm.draw()
            
        except Exception as e:
            ax.text(0.5, 0.5, f"繪圖錯誤: {str(e)}", ha='center')
            self.canvas_hm.draw()

    def setup_stats_tab(self, parent_widget):
        """設定統計摘要"""
        layout = QVBoxLayout(parent_widget)
        txt = QTextEdit()
        txt.setReadOnly(True)
        
        values = np.array([d['value'] for d in self.array_data])
        if len(values) > 0:
            lines = []
            lines.append(f"測量專案：{self.group_name}")
            lines.append("══════════════════════════════")
            lines.append(f"總點數：{len(values)}")
            lines.append("")
            lines.append(f"最大值 (Max)：{values.max():.4f}  (Index: {self.array_data[np.argmax(values)]['index']})")
            lines.append(f"最小值 (Min)：{values.min():.4f}  (Index: {self.array_data[np.argmin(values)]['index']})")
            lines.append(f"峰谷值 (P-V)：{values.max() - values.min():.4f}")
            lines.append("")
            lines.append(f"平均值 (Mean)：{values.mean():.4f}")
            lines.append(f"標準差 (Std) ：{values.std(ddof=1):.4f}")
            
            txt.setPlainText("\n".join(lines))
        else:
            txt.setPlainText("無數據")
            
        layout.addWidget(txt)

