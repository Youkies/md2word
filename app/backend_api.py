# app/backend_api.py
import logging
import os
import json
import platform
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog
import re
import time
import win32com.client
import win32clipboard
import win32con
import pywintypes
import webview

# 关键库导入
import pypandoc
import docx
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

if platform.system() == "Windows":
    import winreg

# 从同级模块导入
from .config import load_config, save_config
from .utils import get_filename_from_content

# 全局变量
window = None  # 这个变量将被主GUI模块注入

# 安全调用window.evaluate_js的辅助函数
def safe_evaluate_js(js_code):
    """安全地调用window.evaluate_js，确保window存在"""
    global window
    if window:
        try:
            window.evaluate_js(js_code)
        except Exception as e:
            logging.error(f"调用JavaScript代码失败: {e}")
    else:
        logging.warning(f"window未初始化，无法执行JavaScript: {js_code}")


def _clear_clipboard():
    """Safely clears the clipboard."""
    clipboard_opened = False
    try:
        win32clipboard.OpenClipboard()
        clipboard_opened = True
        win32clipboard.EmptyClipboard()
        logging.info("Clipboard cleared at start.")
    except pywintypes.error as e:
        # It's not critical if this fails, but log it.
        logging.warning(f"启动时清空剪贴板失败 (可能已被占用): {e}")
    finally:
        if clipboard_opened:
            win32clipboard.CloseClipboard()


def _preprocess_markdown(text):
    """对 Markdown 文本进行预处理，以解决因缺少空行导致的 Pandoc 转换失败问题。"""
    # 步骤 1: 统一进行文本替换
    # 修正 LaTex 公式中的多余反斜杠
    text = text.replace('\\\\', '\\')
    # 采纳您的建议，将 <br> 和 \<br> 替换为两个空格
    text = text.replace('\<br\>', '  ').replace('<br>', '  ')
    # 替换不间断空格
    text = text.replace('\u00A0', ' ').replace('\u3000', ' ')

    # 步骤 2: 智能插入空行，保证块级元素之间格式正确
    lines = text.split('\n')
    processed_lines = []
    for i, current_line in enumerate(lines):
        processed_lines.append(current_line)
        if i < len(lines) - 1:
            next_line = lines[i + 1]
            current_line_stripped = current_line.strip()
            next_line_stripped = next_line.strip()
            if not current_line_stripped or not next_line_stripped:
                continue
            
            # 修复之前损坏的正则表达式
            is_normal_text_line = (not current_line_stripped.startswith(('* ', '- ', '+ ', '>', '#', '|')) and
                                   not re.match(r'^\d+\.\s', current_line_stripped))
            is_next_line_block_start = (next_line_stripped.startswith(('* ', '- ', '+ ', '|', '>')) or
                                        re.match(r'^\d+\.\s', next_line_stripped))
            is_current_line_list_item = (current_line_stripped.startswith(('* ', '- ', '+ ')) or
                                         re.match(r'^\d+\.\s', current_line_stripped))
            is_next_line_table_start = next_line_stripped.startswith('|')
            
            if (is_normal_text_line and is_next_line_block_start) or \
                    (is_current_line_list_item and is_next_line_table_start):
                # 修复之前损坏的正则表达式
                if re.match(r'^[\s|: -]+$', current_line_stripped) and '|' in current_line_stripped:
                    continue
                logging.info(f"自动修正：在第 {i + 1} 行后插入了一个空行以确保块级元素格式正确。")
                processed_lines.append('')
                
    return '\n'.join(processed_lines)


def create_reference_docx(styles):
    """根据传入的样式字典动态创建一个 reference.docx 文件"""
    # ... (此函数代码与原文件相同，复制到此处即可)
    try:
        document = docx.Document()

        def _apply_font_style(style, style_data):
            font = style.font
            font_name = style_data.get('font', '宋体')
            size = float(style_data.get('size', 12))
            color_hex = style_data.get('color', '000000')
            font.size = Pt(size)
            font.color.rgb = RGBColor.from_string(color_hex)
            font.name = font_name
            rfonts = font._element.rPr.rFonts
            rfonts.set(qn('w:eastAsia'), font_name)
            theme_attrs = ['asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme']
            for attr in theme_attrs:
                if rfonts.get(qn(f'w:{attr}')) is not None:
                    rfonts.attrib.pop(qn(f'w:{attr}'))

        _apply_font_style(document.styles['Normal'], styles.get('body', {}))
        for i in range(1, 4):
            _apply_font_style(document.styles[f'Heading {i}'], styles.get(f'h{i}', {}))

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.docx')
        document.save(temp_file.name)
        logging.info(f"动态样式文件已创建: {temp_file.name}")
        return temp_file.name
    except Exception as e:
        logging.error(f"创建动态样式文件失败: {e}", exc_info=True)
        return None


class Api:
    def __init__(self):
        # ... (整个 Api 类的代码与原文件相同，复制到此处即可)
        self.config = load_config()
        self.export_directory = self.config.get('export_directory')
        self.custom_template_path = self.config.get('template_path')
        self.styles = self.config.get('styles')
        self.text_processing = self.config.get('text_processing', {'remove_separators': False})
        self.temp_files_to_clean = []
        self.preset_styles = {
            "general": {"body": {"font": "宋体", "size": 12, "color": "000000"},
                        "h1": {"font": "黑体", "size": 22, "color": "000000"},
                        "h2": {"font": "黑体", "size": 16, "color": "000000"},
                        "h3": {"font": "黑体", "size": 14, "color": "000000"}},
            "academic": {"body": {"font": "Times New Roman", "size": 12, "color": "000000"},
                         "h1": {"font": "Times New Roman", "size": 16, "color": "000000"},
                         "h2": {"font": "Times New Roman", "size": 14, "color": "000000"},
                         "h3": {"font": "Times New Roman", "size": 12, "color": "000000"}},
            "business": {"body": {"font": "Arial", "size": 11, "color": "000000"},
                         "h1": {"font": "Arial", "size": 24, "color": "1F497D"},
                         "h2": {"font": "Arial", "size": 18, "color": "4F81BD"},
                         "h3": {"font": "Arial", "size": 14, "color": "4F81BD"}},
            "technical": {"body": {"font": "Courier New", "size": 11, "color": "000000"},
                          "h1": {"font": "黑体", "size": 18, "color": "000000"},
                          "h2": {"font": "黑体", "size": 16, "color": "000000"},
                          "h3": {"font": "黑体", "size": 14, "color": "000000"}},
            "teaching": {"body": {"font": "楷体", "size": 14, "color": "000000"},
                         "h1": {"font": "黑体", "size": 22, "color": "000000"},
                         "h2": {"font": "黑体", "size": 18, "color": "000000"},
                         "h3": {"font": "楷体", "size": 16, "color": "000000"}},
            "government": {"body": {"font": "仿宋", "size": 16, "color": "000000"},
                           "h1": {"font": "黑体", "size": 22, "color": "000000"},
                           "h2": {"font": "楷体", "size": 16, "color": "000000"},
                           "h3": {"font": "仿宋", "size": 16, "color": "000000"}},
            "modern": {"body": {"font": "微软雅黑", "size": 11, "color": "333333"},
                       "h1": {"font": "微软雅黑", "size": 24, "color": "0078D4"},
                       "h2": {"font": "微软雅黑", "size": 18, "color": "0078D4"},
                       "h3": {"font": "微软雅黑", "size": 15, "color": "333333"}}
        }

    def get_preset_styles(self):
        return self.preset_styles

    def get_system_fonts(self):
        preferred_fonts = ['宋体', '黑体', '楷体', '仿宋', '微软雅黑', 'Times New Roman', 'Arial', 'Courier New']
        all_system_fonts = set()
        try:
            if platform.system() == "Windows":
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as key:
                    i = 0
                    while True:
                        try:
                            name, _, _ = winreg.EnumValue(key, i)
                            all_system_fonts.add(name.split(' (')[0].strip())
                            i += 1
                        except OSError:
                            break
            elif platform.system() == "Darwin":
                output = subprocess.check_output(['system_profiler', 'SPFontsDataType'], text=True, encoding='utf-8')
                for line in output.splitlines():
                    if not line.startswith(' ') and line.endswith(':'):
                        font_name = line[:-1].strip()
                        if font_name and font_name != "Fonts": all_system_fonts.add(font_name)
            else:
                output = subprocess.check_output(['fc-list', ':', 'family'], text=True, encoding='utf-8')
                for line in output.splitlines():
                    font_family = line.split(',')[0].strip()
                    if font_family: all_system_fonts.add(font_family)
        except Exception as e:
            logging.error(f"获取系统字体失败: {e}. 将仅使用常用字体列表。")
            return preferred_fonts
        final_list = list(preferred_fonts)
        added_fonts = set(preferred_fonts)
        for font in sorted(list(all_system_fonts)):
            if font not in added_fonts: final_list.append(font)
        logging.info(f"检测到 {len(all_system_fonts)} 个系统字体。最终字体列表长度 {len(final_list)}。")
        return final_list

    def get_initial_info(self):
        return self.config

    def save_styles(self, styles, last_preset=""):
        self.styles = styles
        self.config['styles'] = styles
        self.config['last_preset'] = last_preset
        save_config(self.config)
        return {"success": True}

    def update_filename(self, content):
        return get_filename_from_content(content)

    def select_export_directory(self):
        global window
        if not window:
            logging.error("Window object not available for folder dialog.")
            return None

        result = window.create_file_dialog(webview.FOLDER_DIALOG, directory=self.export_directory)
        
        if result and result[0]:
            directory = result[0]
            self.export_directory = directory
            self.config['export_directory'] = directory
            save_config(self.config)
            return directory
        return None

    def _run_dialog_in_thread(self, func):
        threading.Thread(target=func, daemon=True).start()

    def open_file_dialog(self):
        global window
        if not window:
            logging.error("Window object not available for file dialog.")
            return

        file_types = ("Markdown 文件 (*.md;*.markdown)", "所有文件 (*.*)")
        result = window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)

        if result and result[0]:
            filepath = result[0]
            logging.info(f"正在打开文件: {filepath}")
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
                # 应用文本处理
                processed_content = self.process_text(content)
                safe_evaluate_js(f'window.app.onFileOpened({json.dumps(processed_content)})')
            except Exception as e:
                logging.error(f"读取文件失败: {filepath}", exc_info=True)
                safe_evaluate_js(f'window.app.showNotification("无法读取文件: {e}", "error")')

    def select_template_dialog(self):
        global window
        if not window:
            logging.error("Window object not available for file dialog.")
            return

        file_types = ("Word 文档 (*.docx)",)
        result = window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=file_types)
        
        if result and result[0]:
            filepath = result[0]
            logging.info(f"用户选择了新模板: {filepath}")
            self.custom_template_path = filepath
            self.config['template_path'] = filepath
            self.config['last_preset'] = 'custom'
            save_config(self.config)
            base_name = os.path.basename(filepath)
            safe_evaluate_js(f'window.app.onTemplateSelected({json.dumps(base_name)})')

    def cleanup_on_exit(self):
        """在程序关闭时，清理所有为复制功能创建的临时文件。"""
        logging.info("程序关闭，开始清理临时文件...")
        for f in self.temp_files_to_clean:
            if f and os.path.exists(f):
                try:
                    os.remove(f)
                    logging.info(f"已删除临时文件: {f}")
                except Exception as e:
                    logging.error(f"删除临时文件失败 {f}: {e}")
        self.temp_files_to_clean.clear()

    def copy_via_office_app(self, content, styles, target_app='word'):
        def _copy():
            # 0. 在开始前清空剪贴板，确保一个干净的环境
            _clear_clipboard()

            # 使用与保存相同的前处理逻辑
            processed_content = _preprocess_markdown(content)
            if not processed_content.strip():
                safe_evaluate_js('window.app.showNotification("内容为空，无法复制。", "info")')
                return

            temp_ref_path = None
            temp_output_path = None
            office_app = None

            try:
                # 1. 创建一个临时的Word文档
                fd, temp_output_path = tempfile.mkstemp(suffix='.docx', prefix='md_copy_')
                os.close(fd)
                # 登记到待清理列表
                self.temp_files_to_clean.append(temp_output_path)

                # 强制使用根目录的 "楷体模板.docx"
                extra_args = ['--mathjax']
                template_to_use = os.path.join(os.getcwd(), '楷体模板.docx')
                if os.path.exists(template_to_use):
                    extra_args.append(f'--reference-doc={template_to_use}')
                    logging.info(f"复制预览功能强制使用模板: {template_to_use}")
                else:
                    logging.warning(f"楷体模板.docx 未找到，将使用Pandoc默认样式进行复制。")
                
                pypandoc.convert_text(source=processed_content, to='docx',
                                      format='markdown+tex_math_dollars+tex_math_single_backslash',
                                      outputfile=temp_output_path, extra_args=extra_args)

                # 2. 在后台启动一个全新的、独立的对应Office实例
                if target_app == 'word':
                    prog_id = "Word.Application"
                    logging.info(f"正在尝试启动后台应用: {prog_id}")
                    office_app = win32com.client.DispatchEx(prog_id)
                else: # wps
                    office_app = None
                    # WPS Office has several possible ProgIDs depending on version and installation.
                    # We try them in order to find one that works.
                    wps_prog_ids = ["wps.application", "kwps.application"]
                    logging.info(f"正在尝试连接 WPS Office, 将依次尝试: {wps_prog_ids}")
                    for prog_id in wps_prog_ids:
                        try:
                            office_app = win32com.client.DispatchEx(prog_id)
                            logging.info(f"成功连接到 WPS Office: {prog_id}")
                            break # Success, exit loop
                        except pywintypes.com_error:
                            logging.warning(f"尝试连接 {prog_id} 失败, 正在尝试下一个...")
                            continue
                    
                    if not office_app:
                        # If loop finishes without success, raise a specific error
                        raise Exception("无法连接到 WPS Office。请确认 WPS 已正确安装并注册了 COM 组件。")

                office_app.Visible = False
                
                # 策略变更：不再直接打开文件，而是新建一个内存文档后插入文件内容，
                # 这样可以彻底避免在"最近文件"列表中留下痕迹。
                doc = office_app.Documents.Add()
                logging.info("创建了一个新的内存文档，用以承载复制内容。")
                
                # 将pandoc生成的临时文件内容插入到新文档中
                doc.Content.InsertFile(FileName=temp_output_path, ConfirmConversions=False, Link=False)
                logging.info(f"已将临时文件 {temp_output_path} 的内容插入内存文档。")
                
                doc.Content.Select()
                doc.Content.Copy()
                time.sleep(0.2) # 增加短暂延时，等待Office完成剪贴板操作
                
                # -- 强制刷新剪贴板，获取所有权 --
                clipboard_opened = False
                try:
                    # Retry opening the clipboard, as it might be locked by the Office app momentarily
                    for i in range(10): # Increased retries
                        try:
                            win32clipboard.OpenClipboard()
                            clipboard_opened = True
                            break
                        except pywintypes.error:
                            time.sleep(0.1)
                    
                    if not clipboard_opened:
                        raise Exception("无法打开剪贴板，它可能正被另一个程序持续占用。")

                    clipboard_data = {}
                    fmt = 0
                    while True:
                        fmt = win32clipboard.EnumClipboardFormats(fmt)
                        if fmt == 0: 
                            break
                        try:
                            clipboard_data[fmt] = win32clipboard.GetClipboardData(fmt)
                        except pywintypes.error:
                            pass
                    
                    # 根据用户反馈，移除会引起冲突的RTF格式
                    rtf_format_id = win32clipboard.RegisterClipboardFormat("Rich Text Format")
                    if rtf_format_id in clipboard_data:
                        del clipboard_data[rtf_format_id]
                        logging.info("已从待写入数据中移除RTF格式。")

                    win32clipboard.EmptyClipboard()
                    
                    for fmt, data in clipboard_data.items():
                        try:
                            win32clipboard.SetClipboardData(fmt, data)
                        except pywintypes.error:
                            pass
                    
                    logging.info(f"剪贴板已强制刷新并接管，保留了 {len(clipboard_data)} 种格式。")
                except Exception as e:
                    logging.error(f"强制刷新剪贴板失败: {e}")
                finally:
                    if clipboard_opened:
                        try:
                            win32clipboard.CloseClipboard()
                        except pywintypes.error as e_close:
                            # 如果关闭失败，很可能是因为它从未被成功打开，或者已经被Office关闭了。
                            # 记录一个警告而不是让程序崩溃。
                            logging.warning(f"关闭剪贴板时发生错误 (可忽略): {e_close}")
                # -- 刷新结束 --

                doc.Close(SaveChanges=False)
                safe_evaluate_js('window.app.showNotification("内容已复制到剪贴板。", "success")')

            except Exception as e:
                error_msg_prefix = f"无法启动后台应用。请确保您已正确安装 {target_app.upper()}"
                if "pywintypes.com_error" in str(e) or "无效的类字符串" in str(e) or "无法连接" in str(e):
                    error_msg = f"{error_msg_prefix} 并且程序有足够权限。"
                else:
                    error_str = str(e).replace('"', "'")
                    error_msg = f"复制失败: {error_str}"
                logging.error(f"通过 {target_app} 复制时失败: {e}", exc_info=True)
                safe_evaluate_js(f'window.app.showNotification("{error_msg}", "error")')
            finally:
                # 3. 清理资源 (只清理Office进程和临时的样式文件)
                if office_app:
                    office_app.Quit()
                if temp_ref_path and os.path.exists(temp_ref_path):
                    os.remove(temp_ref_path)

        self._run_dialog_in_thread(_copy)

    def save_word_document(self, content, directory, filename, styles):
        def _save():
            processed_content = _preprocess_markdown(content)
            if not all([processed_content, directory, filename]):
                safe_evaluate_js('window.app.showNotification("内容、保存路径或文件名不能为空。", "error")')
                return
            output_path = os.path.join(directory, f"{filename}.docx")
            temp_ref_path = None
            try:
                logging.info(f"正在保存文件到: {output_path}")
                extra_args = ['--mathjax']
                if self.custom_template_path and self.config.get('last_preset') == 'custom' and os.path.exists(
                        self.custom_template_path):
                    extra_args.append(f'--reference-doc={self.custom_template_path}')
                    logging.info(f"使用用户选择的模板: {self.custom_template_path}")
                else:
                    logging.info("未选择模板，根据样式设置动态生成...")
                    temp_ref_path = create_reference_docx(styles)
                    if temp_ref_path:
                        extra_args.append(f'--reference-doc={temp_ref_path}')
                    else:
                        logging.warning("动态样式文件创建失败，将使用Pandoc默认样式。")

                pypandoc.convert_text(source=processed_content, to='docx',
                                      format='markdown+tex_math_dollars+tex_math_single_backslash',
                                      outputfile=output_path, extra_args=extra_args)
                safe_path = json.dumps(output_path)
                safe_evaluate_js(f'window.app.showExportSuccessDialog({safe_path})')

            except Exception as e:
                error_str = str(e).replace('"', "'")
                error_msg = f"转换失败: {error_str}"
                logging.error(f"Pandoc DOCX conversion failed: {e}", exc_info=True)
                safe_evaluate_js(f'window.app.showNotification("{error_msg}", "error")')

            finally:
                if temp_ref_path and os.path.exists(temp_ref_path):
                    os.remove(temp_ref_path)
                    logging.info(f"临时样式文件已删除: {temp_ref_path}")

        self._run_dialog_in_thread(_save)

    def get_clipboard_content(self):
        try:
            root = tk.Tk();
            root.withdraw()
            content = root.clipboard_get();
            root.destroy()
            return content
        except Exception:
            return ""

    def open_file(self, path):
        self._run_dialog_in_thread(lambda: self._open_path(path))

    def open_folder(self, path):
        self._run_dialog_in_thread(lambda: self._open_path(os.path.dirname(path)))

    def _open_path(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", path], check=True)
            else:
                subprocess.run(["xdg-open", path], check=True)
        except Exception as e:
            logging.error(f"打开路径失败: {path} - {e}")

    def process_text(self, content):
        """
        根据文本处理设置处理输入的文本
        """
        if not content:
            return ""
            
        processed_text = content
        
        # 删除分隔线
        if self.text_processing.get('remove_separators', False):
            # 匹配任何三个以上的连字符（---）作为分隔线
            processed_text = re.sub(r'^---+$', '', processed_text, flags=re.MULTILINE)
        
        return processed_text
        
    def save_text_processing_settings(self, settings):
        """
        保存文本处理设置
        """
        try:
            self.text_processing = settings
            self.config['text_processing'] = settings
            save_config(self.config)
            logging.info(f"已保存文本处理设置: {settings}")
            return {"success": True}
        except Exception as e:
            logging.error(f"保存文本处理设置失败: {e}")
            return {"success": False, "error": str(e)}
    
    def get_text_processing_settings(self):
        """
        获取当前的文本处理设置
        """
        return self.text_processing