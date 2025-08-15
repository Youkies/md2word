# app/config.py
import os
import json
import logging
import sys
import platform
from .utils import get_app_data_dir, get_default_directory

def get_config_path():
    """获取位于 app data 目录中的配置文件的路径"""
    return os.path.join(get_app_data_dir(), 'config.json')

def save_config(data):
    """保存配置到 JSON 文件"""
    try:
        with open(get_config_path(), 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        logging.info(f"配置已保存: {data}")
    except Exception as e:
        logging.error(f"保存配置文件失败: {e}")

def load_config():
    """从 JSON 文件加载配置，并与默认配置合并"""
    config_path = get_config_path()
    default_config = {
        "export_directory": get_default_directory(),
        "template_path": "",
        "styles": {
            "body": {"font": "宋体", "size": 10.5, "color": "000000"},
            "h1": {"font": "黑体", "size": 18, "color": "000000"},
            "h2": {"font": "黑体", "size": 16, "color": "000000"},
            "h3": {"font": "黑体", "size": 15, "color": "000000"}
        },
        "last_preset": "general",
        "text_processing": {
            "remove_separators": False
        }
    }
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                # 智能合并，防止新版本增加配置项后旧配置文件出错
                for key, value in loaded.items():
                    if isinstance(value, dict) and isinstance(default_config.get(key), dict):
                        default_config[key].update(value)
                    else:
                        default_config[key] = value
                logging.info(f"配置已加载: {default_config}")
                return default_config
        except Exception as e:
            logging.error(f"加载配置文件失败: {e}")
    return default_config

def get_app_data_dir():
    """获取应用数据目录，以便存储日志和配置文件"""
    if platform.system() == "Windows":
        return os.path.join(os.environ['APPDATA'], 'MarkdownConverter')
    else: # macOS and Linux
        return os.path.join(os.path.expanduser('~'), '.config', 'MarkdownConverter')

def setup_logging():
    """配置日志系统"""
    log_dir = get_app_data_dir()
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, 'app.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - [%(threadName)s] - %(message)s',
        handlers=[
            logging.FileHandler(log_file, mode='w', encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

class AppConfig:
    def __init__(self):
        app_data_dir = get_app_data_dir()
        os.makedirs(app_data_dir, exist_ok=True)
        self.config_path = os.path.join(app_data_dir, 'config.json')
        self.config = self._load()

    def _load(self):
        """从 JSON 文件加载配置，并与默认配置合并"""
        default_config = {
            "export_directory": get_default_directory(),
            "template_path": "",
            "styles": {
                "body": {"font": "宋体", "size": 10.5, "color": "000000"},
                "h1": {"font": "黑体", "size": 18, "color": "000000"},
                "h2": {"font": "黑体", "size": 16, "color": "000000"},
                "h3": {"font": "黑体", "size": 15, "color": "000000"}
            },
            "last_preset": "general",
            "text_processing": {
                "remove_separators": False
            }
        }
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    # 智能合并，防止新版本增加配置项后旧配置文件出错
                    for key, value in loaded.items():
                        if isinstance(value, dict) and isinstance(default_config.get(key), dict):
                            default_config[key].update(value)
                        else:
                            default_config[key] = value
                    logging.info(f"配置已加载: {default_config}")
                    return default_config
            except Exception as e:
                logging.error(f"加载配置文件失败: {e}")
        return default_config

    def _save(self):
        """保存配置到 JSON 文件"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            logging.info(f"配置已保存至 {self.config_path}")
        except Exception as e:
            logging.error(f"保存配置文件失败: {e}")

    def get(self, key, default=None):
        """获取配置项"""
        return self.config.get(key, default)

    def set(self, key, value):
        """设置配置项"""
        self.config[key] = value
        self._save()

    def _get_default_config(self):
        """获取默认配置"""
        return {
            "version": "3.0",
            "export_directory": os.path.expanduser('~/Documents'),
            "template_path": None,
            "last_preset": "default",
            "styles": {
                "body": {"font": "宋体", "size": 12, "color": "000000"},
                "h1": {"font": "黑体", "size": 22, "color": "000000"},
                "h2": {"font": "黑体", "size": 18, "color": "000000"},
                "h3": {"font": "黑体", "size": 15, "color": "000000"}
            },
            "text_processing": {
                "remove_separators": False
            }
        }