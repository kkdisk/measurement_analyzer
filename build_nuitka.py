# -*- coding: utf-8 -*-
"""
Measurement Analyzer - Nuitka 打包腳本
使用 Nuitka 編譯為原生執行檔，獲得更小的檔案與更快的啟動速度
"""
import subprocess
import os
import sys
import shutil

# 設定
MAIN_SCRIPT = "main.py"
APP_NAME = "MeasurementAnalyzer"
OUTPUT_DIR = "dist_nuitka"

def get_file_size_mb(filepath):
    """取得檔案大小 (MB)"""
    if os.path.exists(filepath):
        return os.path.getsize(filepath) / (1024 * 1024)
    return 0

def build_with_nuitka():
    print(f"=== Nuitka 打包: {APP_NAME} ===\n")
    
    # 1. 清理舊的輸出資料夾
    if os.path.exists(OUTPUT_DIR):
        try:
            shutil.rmtree(OUTPUT_DIR)
            print(f"已刪除舊資料夾: {OUTPUT_DIR}")
        except Exception as e:
            print(f"警告: 無法刪除 {OUTPUT_DIR}: {e}")
    
    # 2. 建立輸出資料夾
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # 3. Nuitka 編譯參數
    # 注意：使用 standalone 模式（資料夾）而非 onefile（單檔）
    # 因為 onefile 在 Windows 上容易出現資源嵌入失敗的問題
    nuitka_args = [
        sys.executable, "-m", "nuitka",
        
        # 基本設定 - 使用 standalone 模式
        "--standalone",                    # 獨立執行，包含所有依賴
        "--onefile",                     # 單一執行檔 (停用，避免資源嵌入問題)
        #"--output-dir={OUTPUT_DIR}",
        
        # Windows 設定
        "--windows-disable-console",       # 隱藏主控台 (GUI 程式)
        # "--windows-icon-from-ico=app_icon.ico",  # 若有圖示可啟用
        
        # 插件與依賴
        "--enable-plugin=pyqt6",           # PyQt6 支援
        
        # 明確包含的模組 (避免遺漏)
        "--include-module=config",
        "--include-module=parsers", 
        "--include-module=statistics",
        "--include-module=widgets",
        "--include-module=workers",
        "--include-package=pdfplumber",
        "--include-package=scipy",
        "--include-package=scipy.stats",
        "--include-package=natsort",
        "--include-package=matplotlib",
        "--include-package=pandas",
        "--include-package=numpy",
        
        # 排除不需要的模組 (縮小體積)
        "--nofollow-import-to=tkinter",
        "--nofollow-import-to=IPython",
        "--nofollow-import-to=notebook",
        "--nofollow-import-to=dask",
        "--nofollow-import-to=torch",
        "--nofollow-import-to=tensorflow",
        "--nofollow-import-to=tensorboard",
        "--nofollow-import-to=PyQt5",
        "--nofollow-import-to=PySide2",
        "--nofollow-import-to=PySide6",
        
        # 效能優化
        # "--assume-yes-for-downloads",      # 自動下載 C 編譯器 (首次)
        # "--remove-output",               # 保留中間檔案（加速後續編譯）
        
        # 主程式
        MAIN_SCRIPT
    ]
    
    print("正在執行 Nuitka 編譯 (standalone 模式)...")
    print("(首次編譯可能需要 10-20 分鐘，請耐心等候)\n")
    
    # 4. 執行編譯
    try:
        result = subprocess.run(nuitka_args, check=True)
        print(f"\n✅ 編譯成功！")
        
        # 5. 顯示結果 - standalone 模式會產生 main.dist 資料夾
        dist_folder = os.path.join(OUTPUT_DIR, "main.dist")
        exe_path = os.path.join(dist_folder, "main.exe")
        
        if os.path.exists(exe_path):
            size_mb = get_file_size_mb(exe_path)
            print(f"\n📁 輸出資料夾: {os.path.abspath(dist_folder)}")
            print(f"🚀 執行檔: {os.path.abspath(exe_path)}")
            print(f"📊 執行檔大小: {size_mb:.2f} MB")
            
            # 計算整個資料夾大小
            total_size = 0
            for dirpath, dirnames, filenames in os.walk(dist_folder):
                for f in filenames:
                    fp = os.path.join(dirpath, f)
                    total_size += os.path.getsize(fp)
            total_mb = total_size / (1024 * 1024)
            print(f"📦 資料夾總大小: {total_mb:.2f} MB")
            
            # 比較 PyInstaller 版本
            pyinstaller_exe = os.path.join("dist", f"{APP_NAME}.exe")
            if os.path.exists(pyinstaller_exe):
                pi_size = get_file_size_mb(pyinstaller_exe)
                print(f"\n📊 與 PyInstaller 比較:")
                print(f"   PyInstaller: {pi_size:.2f} MB (單檔)")
                print(f"   Nuitka:      {total_mb:.2f} MB (資料夾)")
            
            print(f"\n💡 提示: 可將 main.dist 資料夾重新命名為 {APP_NAME}")
        else:
            # 尋找其他可能的執行檔位置
            print(f"\n⚠️ 在預期位置找不到執行檔，搜尋中...")
            for root, dirs, files in os.walk(OUTPUT_DIR):
                for f in files:
                    if f.endswith(".exe"):
                        exe_found = os.path.join(root, f)
                        size_mb = get_file_size_mb(exe_found)
                        print(f"找到執行檔: {os.path.abspath(exe_found)}")
                        print(f"檔案大小: {size_mb:.2f} MB")
                        break
                        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 編譯失敗: {e}")
        print("請確認已安裝 Nuitka: pip install nuitka")
        return False
    except FileNotFoundError:
        print("\n❌ 錯誤: 找不到 Nuitka，請先安裝:")
        print("   pip install nuitka")
        return False
    
    return True

def main():
    # 檢查 Nuitka 是否已安裝
    try:
        from importlib.metadata import version
        nuitka_version = version("nuitka")
        print(f"Nuitka 版本: {nuitka_version}\n")
    except Exception:
        print("❌ Nuitka 未安裝，正在安裝...")
        subprocess.run([sys.executable, "-m", "pip", "install", "nuitka"], check=True)
        print("✅ Nuitka 安裝完成\n")
    
    build_with_nuitka()

if __name__ == "__main__":
    main()
