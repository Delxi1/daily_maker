import psutil
import time
import sys
import os
import signal
import subprocess
from datetime import datetime

def log(message):
    """记录日志信息"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def is_valid_cmdline(cmdline):
    """检查命令行参数是否有效"""
    return cmdline is not None and all(isinstance(arg, str) for arg in cmdline)

def find_processes(name):
    """查找包含指定名称的所有进程"""
    return [p for p in psutil.process_iter(['name', 'cmdline'])
            if is_valid_cmdline(p.info['cmdline']) and name in ' '.join(p.info['cmdline'])]

def find_view_process():
    """查找View1.py进程"""
    return find_processes('view.py')

def find_dingding_process():
    """查找DingDing.py进程"""
    return find_processes('DingDing.py')

def terminate_processes(processes):
    """终止指定的进程列表"""
    for process in processes:
        try:
            log(f"终止进程 {process.pid}: {process.info['cmdline']}")
            process.terminate()
            # 等待进程终止
            if not process.wait(timeout=3):
                process.kill()
        except Exception as e:
            log(f"终止进程 {process.pid} 时出错: {e}")

def main():
    log("监控程序已启动，正在监听GUI进程...")

    # 启动GUI进程
    try:
        gui_process = subprocess.Popen(['python', 'view.py'])
        log(f"已启动GUI进程，PID: {gui_process.pid}")
    except Exception as e:
        log(f"启动GUI进程时出错: {e}")
        sys.exit(1)

    while True:
        try:
            # 检查GUI进程是否还在运行
            if gui_process.poll() is not None:
                log("检测到GUI进程已退出")

                # 查找并终止所有DingDing.py进程
                dingding_processes = find_dingding_process()
                if dingding_processes:
                    log(f"找到 {len(dingding_processes)} 个DingDing进程")
                    terminate_processes(dingding_processes)
                else:
                    log("未找到DingDing进程")

                # 查找并终止所有View1.py进程(除了监控程序自身)
                view_processes = find_view_process()
                if view_processes:
                    log(f"找到 {len(view_processes)} 个View进程")
                    terminate_processes(view_processes)
                else:
                    log("未找到View进程")

                log("监控程序完成任务，退出")
                sys.exit(0)

            # 休眠一段时间再检查
            time.sleep(1)

        except KeyboardInterrupt:
            log("用户中断，正在清理并退出...")
            # 终止GUI进程(如果还在运行)
            if gui_process.poll() is None:
                gui_process.terminate()
            # 查找并终止所有相关进程
            terminate_processes(find_dingding_process())
            terminate_processes(find_view_process())
            sys.exit(0)

        except Exception as e:
            log(f"监控过程中发生错误: {e}")
            time.sleep(5)  # 发生错误时等待更长时间

if __name__ == "__main__":
    main()