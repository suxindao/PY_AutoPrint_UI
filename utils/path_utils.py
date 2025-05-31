import sys
import os
from pathlib import Path


def get_app_path(*paths):
    """获取应用程序路径，兼容开发模式和打包后模式"""
    if getattr(sys, 'frozen', False):
        # 打包后模式
        base_path = os.path.dirname(sys.executable)
    else:
        # 开发模式
        base_path = os.path.dirname(os.path.abspath(__file__))
        base_path = os.path.dirname(base_path)  # 向上退一级到项目根目录

    return os.path.join(base_path, *paths)


def ensure_directory_exists(dir_path):
    """确保目录存在"""
    os.makedirs(dir_path, exist_ok=True)
    return dir_path
