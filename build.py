import PyInstaller.__main__
import os
import shutil

# 設定主程式檔案名稱
MAIN_SCRIPT = "measurement_analyzer.py"
# 設定產出的 EXE 名稱
APP_NAME = "MeasurementAnalyzer"

def build_exe():
    print(f"=== 開始打包應用程式: {APP_NAME} ===")

    # 1. 清理舊的构建資料夾 (確保乾淨打包)
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"已刪除舊資料夾: {folder}")
            except Exception as e:
                print(f"警告: 無法刪除 {folder} 資料夾: {e}")

    # 2. PyInstaller 參數設定
    # --onefile: 打包成單一 exe 檔
    # --console: 顯示黑色主控台視窗 (除錯用，若程式閃退可以看到錯誤訊息)
    # --name: 指定執行檔名稱
    # --exclude-module: 排除不必要的套件以避免衝突或縮小體積
    params = [
        MAIN_SCRIPT,
        f'--name={APP_NAME}',
        '--windowed',   # 隱藏主控台視窗 (正式發布用)
        #'--console',    # 開啟主控台模式 (除錯用)
        '--onefile',
        '--clean',
        '--noconfirm',
        # '--icon=app_icon.ico', # 若有 icon 可啟用
        
        # --- 排除不必要的模組 ---
        '--exclude-module=tkinter',
        '--exclude-module=scipy',
        '--exclude-module=IPython',
        '--exclude-module=notebook',
        '--exclude-module=dask',
        
        # 排除其他 Qt 綁定以避免衝突
        '--exclude-module=PyQt5',
        '--exclude-module=PySide2',
        '--exclude-module=PySide6',
    ]

    print(f"正在執行 PyInstaller，參數: {params}")

    # 3. 執行打包
    try:
        PyInstaller.__main__.run(params)
        print(f"\n✅ 打包成功！")
        exe_path = os.path.abspath(os.path.join('dist', APP_NAME + '.exe'))
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"執行檔位於: {exe_path}")
            print(f"檔案大小: {size_mb:.2f} MB")
            #print(f"提示: 若執行時發生錯誤，請查看黑色視窗中的錯誤訊息。")
    except Exception as e:
        print(f"\n❌ 打包失敗: {e}")

if __name__ == "__main__":
    build_exe()