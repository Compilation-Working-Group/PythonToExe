"""
配置文件模块
"""

import json
import os
from pathlib import Path
from typing import Dict, Any

CONFIG_FILE = Path("config/settings.json")
DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",
    "temperature": 0.7,
    "max_tokens": 4000,
    "language": "zh-CN",
    "auto_save": True,
    "save_path": "output",
    "theme": "light",
    "font_size": 12,
    "recent_files": []
}

def load_config() -> Dict[str, Any]:
    """
    加载配置文件
    
    Returns:
        配置字典
    """
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 合并默认配置
                for key, value in DEFAULT_CONFIG.items():
                    if key not in config:
                        config[key] = value
                return config
    except Exception as e:
        print(f"加载配置文件失败: {e}")
    
    return DEFAULT_CONFIG.copy()

def save_config(config: Dict[str, Any]):
    """
    保存配置文件
    
    Args:
        config: 配置字典
    """
    try:
        # 确保目录存在
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置文件失败: {e}")

def get_api_key() -> str:
    """
    获取API密钥
    
    Returns:
        API密钥字符串
    """
    config = load_config()
    api_key = config.get("api_key", "")
    
    # 如果没有配置，尝试从环境变量获取
    if not api_key:
        api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    
    return api_key

def update_config(key: str, value: Any):
    """
    更新配置项
    
    Args:
        key: 配置键
        value: 配置值
    """
    config = load_config()
    config[key] = value
    save_config(config)
