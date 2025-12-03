import PyInstaller.__main__
import os
import shutil

# 設定主程式檔案名稱
MAIN_SCRIPT = "measurement_analyzer.py"
# 設定產出的 EXE 名稱
APP_NAME = "MeasurementAnalyzer"

def build_exe():
    print("=== 開始打包應用程式 ===")

    # 1. 清理舊的构建資料夾 (確保乾淨打包)
    if os.path.exists("build"):
        try:
            shutil.rmtree("build")
        except Exception as e:
            print(f"警告: 無法刪除 build 資料夾: {e}")

    if os.path.exists("dist"):
        try:
            shutil.rmtree("dist")
        except Exception as e:
            print(f"警告: 無法刪除 dist 資料夾: {e}")

    # 2. PyInstaller 參數設定
    # --onefile: 打包成單一 exe 檔
    # --windowed: 執行時不顯示黑色主控台視窗 (Console)
    # --name: 指定執行檔名稱
    # --exclude-module: 排除不必要的套件以避免衝突或縮小體積
    params = [
        MAIN_SCRIPT,
        f'--name={APP_NAME}',
        '--windowed',
        '--onefile',
        '--clean',
        '--noconfirm',
        # 若之後有 icon 圖示，可以取消註解下一行並放入 icon.ico
        # '--icon=icon.ico',
        
        # 排除不需要的模組以減小體積 (可選)
        '--exclude-module=tkinter',
        
        # [重要修正] 排除 PyQt5 以解決與 PyQt6 的衝突 (Multiple Qt bindings error)
        '--exclude-module=PyQt5',
        # 為了保險起見，同時排除其他可能導致衝突的 Qt bindings
        '--exclude-module=PySide2',
        '--exclude-module=PySide6',
    ]

    print(f"正在執行 PyInstaller，參數: {params}")

    # 3. 執行打包
    try:
        PyInstaller.__main__.run(params)
        print(f"\n✅ 打包成功！")
        print(f"執行檔位於: {os.path.abspath(os.path.join('dist', APP_NAME + '.exe'))}")
    except Exception as e:
        print(f"\n❌ 打包失敗: {e}")

if __name__ == "__main__":
    build_exe()