# src/__init__.py
"""
Academic Writer Pro - 智能文稿撰写助手
=====================================

一个基于 DeepSeek API 的智能文稿撰写软件，
支持期刊论文、研究计划、反思报告、案例分析、总结报告等多种文档类型。

使用方法:
    from src import run_app
    run_app()

或从命令行:
    python -m src
"""

__version__ = '1.0.0'
__author__ = 'Academic Writer Team'
__license__ = 'MIT'
__url__ = 'https://github.com/yourusername/academic-writer'

import sys
import os
from pathlib import Path

# 获取包路径
PACKAGE_DIR = Path(__file__).parent
PROJECT_ROOT = PACKAGE_DIR.parent

# 确保必要的目录存在
def ensure_directories():
    """确保必要的目录存在"""
    dirs = ['output', 'templates', 'config', 'assets', 'logs']
    for dir_name in dirs:
        (PROJECT_ROOT / dir_name).mkdir(exist_ok=True)

# 初始化
ensure_directories()

# 导出主要函数和类
def run_app():
    """运行应用程序的主函数"""
    from .main import main
    return main()

# 导出主要模块
__all__ = [
    'run_app',
    'PACKAGE_DIR',
    'PROJECT_ROOT',
]

# 当直接运行此包时启动应用程序
if __name__ == '__main__':
    run_app()
