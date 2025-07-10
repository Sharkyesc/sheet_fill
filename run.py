import psutil
import subprocess
import time
import json
import os
import sys
from monitor import SystemMonitor

def get_baseline():
    cpu = psutil.cpu_percent(interval=1)
    mem = psutil.virtual_memory()
    return {
        'cpu_percent': cpu,
        'memory_percent': mem.percent,
        'memory_used_mb': mem.used / (1024 * 1024),
        'memory_available_mb': mem.available / (1024 * 1024)
    }

def main():
    baseline = get_baseline()
    print("baseline:", baseline)

    monitor = SystemMonitor(interval=1000, max_records=3600)
    monitor.start_monitoring()
    time.sleep(2)

    main_args = sys.argv[1:]
    print("开始运行 main.py ... 参数:", main_args)
    subprocess.run(['python', 'main.py'] + main_args)

    monitor.stop_monitoring()
    monitor.print_summary()
    data_path = monitor.save_data()
    chart_path = monitor.generate_charts()

    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    data['memory_used_mb'] = [v - baseline['memory_used_mb'] for v in data['memory_used_mb']]
    data['memory_available_mb'] = [v - baseline['memory_available_mb'] for v in data['memory_available_mb']]
    data['memory_percent'] = [v - baseline['memory_percent'] for v in data['memory_percent']]

    fixed_path = data_path.replace('.json', '_fixed.json')
    with open(fixed_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print("数据已保存：", fixed_path)

    monitor.load_data(fixed_path)
    fixed_chart_path = chart_path.replace('.png', '_fixed.png')
    monitor.generate_charts(fixed_chart_path)
    print("图表已保存：", fixed_chart_path)

if __name__ == '__main__':
    main() 