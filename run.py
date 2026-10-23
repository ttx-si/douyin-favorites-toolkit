"""
一键启动脚本
使用方法：直接在当前目录执行 python run.py
"""
import sys
from pipeline import execute_pipeline

if __name__ == "__main__":
    print("==================================================")
    print(" 抖音收藏夹全自动提取与逐字稿归档工具")
    print("==================================================")
    print("正在连接 Chrome 调试端口并嗅探抖音账号...")
    try:
        execute_pipeline()
    except Exception as e:
        print(f"\n[错误] 流水线执行中断: {e}")
        sys.exit(1)
