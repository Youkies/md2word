# main.py
import os
import sys
import logging
import platform
import ctypes


def main():
    # --- 关键修复：确保程序能找到外部资源（如图片）---
    # 这一步在新结构中至关重要，它能让相对路径正确工作
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # 导入我们自己的模块。必须在 chdir 之后进行，以确保模块能被找到。
    from app.utils import get_app_data_dir, handle_exception
    from app.gui_manager import create_and_run_gui

    # --- 全局异常捕获 ---
    sys.excepthook = handle_exception

    # --- 日志系统设置 ---
    app_data_directory = get_app_data_dir()
    log_file_path = os.path.join(app_data_directory, 'converter_unified.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(threadName)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[
            logging.FileHandler(log_file_path, 'w', 'utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("---------- 应用启动 ----------")

    # --- PyInstaller 打包相关设置 ---
    if getattr(sys, 'frozen', False):
        logging.info("在打包模式下运行。")
        # 当使用PyInstaller打包时，sys._MEIPASS包含打包后的临时目录路径
        try:
            bundle_dir = sys._MEIPASS  # PyInstaller创建的临时目录
            pandoc_path_in_bundle = os.path.join(bundle_dir, 'pandoc', 'pandoc.exe')
            if os.path.exists(pandoc_path_in_bundle):
                # 将打包的 pandoc 路径添加到环境变量，pypandoc 会自动使用它
                os.environ['PYPANDOC_PANDOC'] = pandoc_path_in_bundle
                logging.info(f"Pandoc in bundle found and configured at: {pandoc_path_in_bundle}")
            else:
                logging.error(f"在打包目录中未找到 pandoc.exe: {pandoc_path_in_bundle}")
        except AttributeError:
            logging.error("无法访问 sys._MEIPASS，这可能意味着程序不是通过PyInstaller打包运行的")

    # --- 启动应用 ---
    try:
        # 导入并运行GUI
        from app.gui_manager import create_and_run_gui
        create_and_run_gui()
    except Exception as e:
        logging.critical("在 main 函数中发生致命错误: ", exc_info=True)
        if platform.system() == "Windows":
            ctypes.windll.user32.MessageBoxW(None,
                                             f"程序遇到致命错误: \n{e}\n\n详情请查看 app data 目录中的 error.log 文件。",
                                             "严重错误", 0x10)
        # 退出前确保日志被写入
        logging.shutdown()
        sys.exit(1)

    logging.info("---------- 应用关闭 ----------")


if __name__ == '__main__':
    # 为了解决在多进程环境（尤其是在macOS和Linux上打包后）可能出现的问题
    # 推荐使用 multiprocessing.freeze_support()
    from multiprocessing import freeze_support

    freeze_support()

    main()