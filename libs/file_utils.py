#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File Utils - EZTeach Skill
文件操作工具类
"""

import os
import json
import hashlib
from typing import Dict, List, Any, Optional
from pathlib import Path
from datetime import datetime


class FileUtils:
    """文件操作工具类"""
    
    @staticmethod
    def ensure_dir(path: str) -> Path:
        """确保目录存在"""
        p = Path(path)
        p.mkdir(parents=True, exist_ok=True)
        return p
    
    @staticmethod
    def safe_filename(filename: str) -> str:
        """生成安全的文件名"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename
    
    @staticmethod
    def get_file_info(path: str) -> Dict[str, Any]:
        """获取文件信息"""
        p = Path(path)
        if not p.exists():
            return {'exists': False}
        
        stat = p.stat()
        return {
            'exists': True,
            'name': p.name,
            'suffix': p.suffix,
            'size': stat.st_size,
            'size_mb': round(stat.st_size / 1024 / 1024, 2),
            'created': datetime.fromtimestamp(stat.st_ctime).isoformat(),
            'modified': datetime.fromtimestamp(stat.st_mtime).isoformat()
        }
    
    @staticmethod
    def load_json(path: str, default: Any = None) -> Any:
        """加载JSON文件"""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default
    
    @staticmethod
    def save_json(data: Any, path: str, indent: int = 2) -> None:
        """保存JSON文件"""
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=indent)
    
    @staticmethod
    def file_md5(path: str) -> str:
        """计算文件MD5"""
        hash_md5 = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    
    @staticmethod
    def find_files(directory: str, pattern: str = '*') -> List[str]:
        """查找文件"""
        return [str(p) for p in Path(directory).glob(pattern)]
    
    @staticmethod
    def cleanup_dir(directory: str, pattern: str = '*', older_than: int = 3600) -> int:
        """清理目录中的旧文件"""
        count = 0
        now = datetime.now().timestamp()
        
        for p in Path(directory).glob(pattern):
            if p.is_file():
                age = now - p.stat().st_mtime
                if age > older_than:
                    p.unlink()
                    count += 1
        return count
