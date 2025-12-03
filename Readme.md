# **Keyence CSV 量測數據分析工具 (Measurement Data Analyzer)**

這是一個專為分析 Keyence LM-X 系列或其他類似量測儀器輸出的 CSV 報告而設計的 Python 桌面應用程式。  
它能夠批次讀取多個 CSV 檔案，自動解析量測數據，計算 CPK、不良率，並提供視覺化圖表分析。

## **主要功能**

1. **批次匯入**：  
   * 支援選擇資料夾自動讀取所有 CSV 檔案。  
   * 支援「累加模式」，可多次匯入不同日期的資料夾進行合併分析。  
2. **自動解析**：  
   * 智慧跳過 CSV 檔頭的儀器資訊，自動定位「No」、「實測值」、「設計值」等關鍵欄位。  
   * 自動解析測量日期與時間。  
3. **公差判定**：  
   * 依據設計值與上下限公差自動判定 Pass/Fail。  
   * 設計值為 0 的項目自動略過判定。  
   * 視覺化標示：Fail 項目在表格中以紅色高亮顯示。  
4. **統計分析 (Pro)**：  
   * 計算每個測項的 **樣本數** (總晶片數)、**NG 次數**、**不良率 (%)**。  
   * 計算 **CPK (製程能力指標)**，並依據數值給予顏色燈號 (紅/黃/綠) 提示。  
   * 支援匯出統計報表為 CSV。  
5. **視覺化圖表**：  
   * **分佈直方圖 (Histogram)**：顯示測值分佈與規格界限 (USL/LSL)。  
   * **趨勢圖 (Trend Chart)**：依照測量時間或檔案讀取順序顯示數值變化，協助監控製程飄移。

## **安裝需求**

請確保您的電腦已安裝 Python 3.8 或以上版本。

### **1\. 下載專案**

git clone \[https://github.com/yourusername/measurement-analyzer.git\](https://github.com/yourusername/measurement-analyzer.git)  
cd measurement-analyzer

### **2\. 安裝依賴套件**

pip install \-r requirements.txt

## **使用方法**

### **執行程式**

在終端機 (Terminal) 中執行以下指令啟動應用程式：

python measurement\_analyzer.py

### **操作流程**

1. 點擊 **「1. 加入資料夾」**，選擇包含量測 CSV 檔的資料夾 (可重複操作加入多個資料夾)。  
2. 程式會自動讀取並顯示數據列表。  
3. 若要快速檢查不良品，勾選 **「僅顯示 FAIL 項目」**。  
4. 點擊 **「3. 統計報告」** 查看 CPK 與不良率統計表。  
5. 在主表格中點選任意一列數據，再點擊 **「4. 視覺化選定測項」** 查看該測項的分佈與趨勢圖。  
6. 點擊 **「2. 匯出總表」** 保存整理後的完整數據。

## **打包成執行檔 (.exe)**

若需要在沒有安裝 Python 的 Windows 電腦上執行，可以使用隨附的 build.py 腳本將程式打包成 EXE 檔。

1. 執行打包腳本：  
   python build.py

2. 打包完成後，執行檔將位於 dist/MeasurementAnalyzer.exe。

## **授權**

MIT License