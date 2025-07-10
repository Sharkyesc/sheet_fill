import os
import time
import json
import psutil
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib import rcParams
import numpy as np

rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
rcParams['axes.unicode_minus'] = False

class SystemMonitor:
    """System resource monitor"""
    
    def __init__(self, interval: int = 100, max_records: int = 1000):
        self.interval = interval
        self.max_records = max_records
        self.monitoring = False
        self.monitor_thread = None
        self.data = {
            'timestamps': [],
            'cpu_percent': [],
            'memory_percent': [],
            'memory_used_mb': [],
            'memory_available_mb': []
        }
        self.lock = threading.Lock()
        
        self.output_dir = os.path.join(os.path.dirname(__file__), 'monitor_output')
        os.makedirs(self.output_dir, exist_ok=True)
        
    def get_system_info(self) -> Dict:
        """Get current system information"""
        try:
            cpu_percent = psutil.cpu_percent(interval=0.1)
            
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            memory_used_mb = memory.used / (1024 * 1024)
            memory_available_mb = memory.available / (1024 * 1024)
            
            return {
                'timestamp': datetime.now(),
                'cpu_percent': cpu_percent,
                'memory_percent': memory_percent,
                'memory_used_mb': memory_used_mb,
                'memory_available_mb': memory_available_mb
            }
        except Exception as e:
            print(f"Failed to get system information: {e}")
            return None
    
    def collect_data(self):
        """Data collection thread function"""
        while self.monitoring:
            try:
                system_info = self.get_system_info()
                if system_info:
                    with self.lock:
                        # Add data
                        self.data['timestamps'].append(system_info['timestamp'])
                        self.data['cpu_percent'].append(system_info['cpu_percent'])
                        self.data['memory_percent'].append(system_info['memory_percent'])
                        self.data['memory_used_mb'].append(system_info['memory_used_mb'])
                        self.data['memory_available_mb'].append(system_info['memory_available_mb'])
                        
                        # Limit record count
                        if len(self.data['timestamps']) > self.max_records:
                            for key in self.data:
                                if isinstance(self.data[key], list):
                                    self.data[key] = self.data[key][-self.max_records:]
                
                time.sleep(self.interval / 1000.0)
            except Exception as e:
                print(f"Data collection error: {e}")
                time.sleep(self.interval / 1000.0)
    
    def start_monitoring(self):
        """Start monitoring"""
        if not self.monitoring:
            self.monitoring = True
            self.monitor_thread = threading.Thread(target=self.collect_data, daemon=True)
            self.monitor_thread.start()
            print(f"System monitoring started, interval: {self.interval} ms")
    
    def stop_monitoring(self):
        """Stop monitoring"""
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=2)
        print("System monitoring stopped")
    
    def get_current_stats(self) -> Dict:
        """Get current statistics"""
        with self.lock:
            if not self.data['timestamps']:
                return {}
            
            latest_data = {
                'cpu_percent': self.data['cpu_percent'][-1] if self.data['cpu_percent'] else 0,
                'memory_percent': self.data['memory_percent'][-1] if self.data['memory_percent'] else 0,
                'memory_used_mb': self.data['memory_used_mb'][-1] if self.data['memory_used_mb'] else 0,
                'memory_available_mb': self.data['memory_available_mb'][-1] if self.data['memory_available_mb'] else 0,
                'total_records': len(self.data['timestamps'])
            }
            
            # Calculate statistics
            if len(self.data['cpu_percent']) > 1:
                latest_data.update({
                    'cpu_avg': np.mean(self.data['cpu_percent']),
                    'cpu_max': np.max(self.data['cpu_percent']),
                    'memory_avg': np.mean(self.data['memory_percent']),
                    'memory_max': np.max(self.data['memory_percent'])
                })
            
            return latest_data
    
    def save_data(self, filename: Optional[str] = None) -> str:
        """Save monitoring data to JSON file"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"monitor_data_{timestamp}.json"
        
        filepath = os.path.join(self.output_dir, filename)
        
        with self.lock:
            # Convert datetime objects to strings
            data_to_save = {}
            for key, value in self.data.items():
                if key == 'timestamps':
                    data_to_save[key] = [ts.isoformat() for ts in value]
                else:
                    data_to_save[key] = value
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data_to_save, f, ensure_ascii=False, indent=2)
        
        print(f"Monitoring data saved to: {filepath}")
        return filepath
    
    def load_data(self, filepath: str):
        """Load monitoring data from JSON file"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)
            
            with self.lock:
                # Convert strings back to datetime objects
                self.data = {}
                for key, value in loaded_data.items():
                    if key == 'timestamps':
                        self.data[key] = [datetime.fromisoformat(ts) for ts in value]
                    else:
                        self.data[key] = value
            
            print(f"Monitoring data loaded from {filepath}")
        except Exception as e:
            print(f"Failed to load data: {e}")
    
    def generate_charts(self, save_path: Optional[str] = None) -> str:
        """Generate monitoring charts"""
        if not self.data['timestamps']:
            print("No data available for chart generation")
            return ""
        
        if not save_path:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = os.path.join(self.output_dir, f"monitor_charts_{timestamp}.png")
        
        with self.lock:
            timestamps = self.data['timestamps']
            cpu_data = self.data['cpu_percent']
            memory_data = self.data['memory_percent']
            memory_used = self.data['memory_used_mb']
            memory_available = self.data['memory_available_mb']
        
        if len(timestamps) > 0:
            start_time = timestamps[0]
            seconds = [(ts - start_time).total_seconds() for ts in timestamps]
        else:
            seconds = []
        
        # Create chart
        fig, axes = plt.subplots(1, 3, figsize=(18, 6))
        fig.suptitle('System Resource Monitoring Report', fontsize=16, fontweight='bold')
        
        # 1. CPU usage
        axes[0].plot(seconds, cpu_data, 'b-', linewidth=2, label='CPU Usage')
        axes[0].set_title('CPU Usage (%)')
        axes[0].set_ylabel('Usage (%)')
        axes[0].set_xlabel('Running time (s)')
        axes[0].grid(True, alpha=0.3)
        axes[0].legend()
        
        # 2. Memory usage
        axes[1].plot(seconds, memory_data, 'r-', linewidth=2, label='Memory Usage')
        axes[1].set_title('Memory Usage (%)')
        axes[1].set_ylabel('Usage (%)')
        axes[1].set_xlabel('Running time (s)')
        axes[1].grid(True, alpha=0.3)
        axes[1].legend()
        
        # 3. Memory usage (MB)
        axes[2].plot(seconds, memory_used, 'g-', linewidth=2, label='Used Memory')
        axes[2].plot(seconds, memory_available, 'orange', linewidth=2, label='Available Memory')
        axes[2].set_title('Memory Usage (MB)')
        axes[2].set_ylabel('Memory (MB)')
        axes[2].set_xlabel('Running time (s)')
        axes[2].grid(True, alpha=0.3)
        axes[2].legend()
        
        plt.tight_layout()
        
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Monitoring chart saved to: {save_path}")
        return save_path
    
    def print_summary(self):
        """Print monitoring summary"""
        stats = self.get_current_stats()
        if not stats:
            print("No monitoring data available")
            return
        
        print("\n" + "="*50)
        print("System Monitoring Summary")
        print("="*50)
        print(f"Total records: {stats.get('total_records', 0)}")
        print(f"Current CPU usage: {stats.get('cpu_percent', 0):.2f}%")
        print(f"Current memory usage: {stats.get('memory_percent', 0):.2f}%")
        print(f"Current memory usage: {stats.get('memory_used_mb', 0):.1f} MB")
        print(f"Current available memory: {stats.get('memory_available_mb', 0):.1f} MB")
        
        if 'cpu_avg' in stats:
            print(f"CPU average usage: {stats['cpu_avg']:.2f}%")
            print(f"CPU maximum usage: {stats['cpu_max']:.2f}%")
            print(f"Memory average usage: {stats['memory_avg']:.2f}%")
            print(f"Memory maximum usage: {stats['memory_max']:.2f}%")
        print("="*50)


def main():
    """Main function - for standalone monitoring"""
    import argparse
    
    parser = argparse.ArgumentParser(description='System Resource Monitor')
    parser.add_argument('--interval', type=int, default=100, help='Monitoring interval (milliseconds)')
    parser.add_argument('--max-records', type=int, default=1000, help='Maximum number of records')
    
    args = parser.parse_args()
    
    monitor = SystemMonitor(interval=args.interval, max_records=args.max_records)
    
    try:
        print("Starting system monitoring...")
        print("Press Ctrl+C to stop monitoring")
        monitor.start_monitoring()
        
        # Continuous monitoring until user interrupts
        while True:
            time.sleep(1)
        
    except KeyboardInterrupt:
        print("\nUser interrupted monitoring")
        monitor.stop_monitoring()
        
        # Generate report
        monitor.print_summary()
        monitor.save_data()
        monitor.generate_charts()
        
        print("Monitoring completed!")


if __name__ == "__main__":
    main() 