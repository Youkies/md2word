# app/utils.py
import logging
import os
import sys
import traceback
import platform
import ctypes
from ctypes import wintypes
import re
import base64
import subprocess


def get_app_data_dir():
    """获取跨平台的应用数据目录，并确保其存在。"""
    app_name = "MarkdownToWordConverter"
    if platform.system() == "Windows":
        base_dir = os.environ.get('LOCALAPPDATA', os.path.expanduser("~"))
    elif platform.system() == "Darwin":
        base_dir = os.path.expanduser('~/Library/Application Support')
    else:
        base_dir = os.path.expanduser('~/.local/share')

    app_data_dir = os.path.join(base_dir, app_name)

    try:
        os.makedirs(app_data_dir, exist_ok=True)
    except Exception as e:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"警告: 无法创建应用数据目录 {app_data_dir} ({e})。将使用当前目录。")
        return os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else script_dir

    return app_data_dir


def handle_exception(exc_type, exc_value, exc_traceback):
    """捕获所有未处理的异常，并将其写入 app data 目录中的 error.log 文件。"""
    log_dir = get_app_data_dir()
    error_log_path = os.path.join(log_dir, 'error.log')
    error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    try:
        with open(error_log_path, 'w', encoding='utf-8') as f:
            f.write("A fatal error occurred:\n")
            f.write(error_msg)
    except Exception as e:
        print(f"无法写入致命错误日志: {e}")
    sys.__excepthook__(exc_type, exc_value, exc_traceback)


def get_default_directory():
    """获取Windows下的'文档'目录或通用的用户主目录"""
    try:
        if platform.system() == "Windows":
            CSIDL_PERSONAL = 5
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, 0, buf)
            if os.path.isdir(buf.value):
                return buf.value
    except Exception as e:
        logging.warning(f"无法获取 '文档' 文件夹路径, 错误: {e}")
    return os.path.expanduser("~")


def get_filename_from_content(content):
    """从Markdown内容中提取第一个标题作为文件名"""
    if not content: return "无标题"
    lines = content.splitlines()
    for h_level in range(1, 7):
        prefix = '#' * h_level + ' '
        for line in lines:
            line = line.strip()
            if line.startswith(prefix):
                sanitized_filename = re.sub(r'[\\/*?:"<>|]', "", line[h_level:].strip())
                return sanitized_filename[:50]
    return "无标题"


def image_to_base64(file_path):
    """将图片文件转换为Base64编码的字符串"""
    try:
        # 使用绝对路径来查找图片
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(sys.modules['__main__'].__file__)))
        abs_path = os.path.join(base_path, file_path)
        with open(abs_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        return f"data:image/png;base64,{encoded_string}"
    except (FileNotFoundError, AttributeError):
        logging.error(f"资源图片未找到: {file_path}")
        # 返回一个表示“未找到”的SVG图像
        return "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxNDAiIGhlaWdodD0iMTQwIiB2aWV3Qm94PSIwIDAgMTQwIDE0MCI+PHJlY3Qgd2lkdGg9IjE0MCIgaGVpZ2h0PSIxNDAiIGZpbGw9IiNlZWVlZWUiLz48dGV4dCB4PSI1MCUiIHk9IjUwJSIgZm9udC1mYW1pbHk9InNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTYiIGZpbGw9IiNhYWFhYWEiIHRleHQtYW5jaGyPSJtaWRkbGUiIGR5PSIuM2VtIj5Ob3QgRm91bmQ8L3RleHQ+PC9zdmc+"


def set_dark_title_bar(hwnd: int):
    """为Windows窗口设置暗色标题栏"""
    if sys.platform != 'win32': return
    try:
        import darkdetect
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(value),
                                                   ctypes.sizeof(value))
        logging.info(f"已为窗口句柄 {hwnd} 应用暗色标题栏。")
    except Exception as e:
        logging.error(f"设置暗色标题栏失败: {e}")