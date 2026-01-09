# 工作日報 (Daily Report) - 2026/01/09

## ✅ 今日完成事項

### 1. [v2.3.0] 公差反推功能實作
*   **核心算法**：實作 `calculate_tolerance_for_yield`，支援常態分佈假設下的公差反推 (Z-score)。
*   **UI 整合**：
    *   **統計表**：新增「建議公差(90%)」欄位，提供視覺化顏色標記 (淺紅/淺綠)。
    *   **詳細分析**：新增「📐 公差建議」分頁，支援動態調整目標良率 (80% ~ 99.73%)。
*   **依賴更新**：引入 `scipy` 庫以進行精確統計計算。

### 2. [v2.4.0] 程式碼重構 (Refactoring)
*   **模組化拆分**：將原 1300+ 行的單一檔案 `measurement_analyzer.py` 拆解為模組化架構：
    *   `main.py`: 程式入口與 UI 主視窗
    *   `config.py`: 設定常數 (AppConfig)
    *   `statistics.py`: 統計運算 (CPK, Tolerance)
    *   `parsers.py`: 資料解析 (CSV, PDF)
    *   `widgets.py`: 自定義 UI 元件 (Dialogs, TableItems)
    *   `workers.py`: 背景執行緒 (FileLoader)
*   **問題修復**：解決拆分過程中的循環引用與 imports 遺漏問題。

### 3. 開發環境與部屬
*   **Git 優化**：更新 `.gitignore`，將 `build/`, `dist/` 等產物移除版本控制。
*   **打包流程**：更新 `build.py`：
    *   修正 `sys.setrecursionlimit` 以解決打包 `scipy` 時的遞迴錯誤。
    *   排除不必要的 ML 庫 (torch, tensorflow) 以縮減體積。
    *   成功打包單一執行檔 `MeasurementAnalyzer.exe` (v2.4.0)。

---

## 📅 後續規劃 (ToDo List)

### 短期目標 (v2.5.0 - 錯誤處理與體驗優化)
- [ ] **錯誤處理強化**：
    - 全面檢視並替換裸 `except:` 為具體異常捕獲。
    - 導入更完整的 logging 機制 (如: 錯誤發生時彈出視窗提示)。
- [ ] **資料載入優化**：
    - 大量檔案 (>100) 載入時的確認提示。
    - 表格渲染優化 (減少重繪)。

### 中期目標 (v2.6.0 - 使用者偏好記憶)
- [ ] **設定持久化**：
    - 記住視窗位置與大小。
    - 記住上次使用的目標良率設定。
    - 記住表格欄位寬度調整。

### 長期目標 (品質保證)
- [ ] **單元測試**：為 `statistics.py` 和 `parsers.py` 建立測試案例 (`tests/`)。
- [ ] **文件更新**：更新 `README.md` 架構說明與開發指南。
