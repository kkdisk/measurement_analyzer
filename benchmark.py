import time
import os
import sys
import psutil
import pandas as pd
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QEventLoop

# Add current directory to path
sys.path.append(os.getcwd())

from measurement_analyzer import FileLoaderThread, AppConfig

def run_benchmark():
    print(f"Starting Benchmark for {AppConfig.TITLE}")
    
    # Create dummy files for testing
    test_dir = "benchmark_data"
    os.makedirs(test_dir, exist_ok=True)
    
    print("Generating 100 test files...")
    for i in range(100):
        df = pd.DataFrame({
            'No': range(1, 101),
            '測量專案': [f'Item_{j}' for j in range(1, 101)],
            '實測值': [10.0 + (j%5)*0.1 for j in range(1, 101)],
            '設計值': [10.0] * 100,
            '上限公差': [0.5] * 100,
            '下限公差': [-0.5] * 100
        })
        df.to_csv(os.path.join(test_dir, f"test_{i}.csv"), index=False, encoding='utf-8-sig')
        
    files = [os.path.join(test_dir, f) for f in os.listdir(test_dir) if f.endswith('.csv')]
    
    app = QApplication(sys.argv)
    
    process = psutil.Process(os.getpid())
    mem_before = process.memory_info().rss / 1024 / 1024
    
    start_time = time.time()
    
    loader = FileLoaderThread(files)
    
    # Use EventLoop to wait for thread
    loop = QEventLoop()
    loader.data_loaded.connect(lambda: loop.quit())
    loader.start()
    loop.exec()
    
    end_time = time.time()
    mem_after = process.memory_info().rss / 1024 / 1024
    
    duration = end_time - start_time
    mem_diff = mem_after - mem_before
    
    print(f"\nBenchmark Results:")
    print(f"Time: {duration:.2f} seconds (Target: < 30s)")
    print(f"Memory Increase: {mem_diff:.2f} MB (Target: < 500MB)")
    
    # Cleanup
    for f in files:
        os.remove(f)
    os.rmdir(test_dir)
    
    if duration < 30 and mem_diff < 500:
        print("PASS: Performance is within acceptable limits.")
    else:
        print("FAIL: Performance targets exceeded.")

if __name__ == "__main__":
    run_benchmark()
