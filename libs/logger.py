#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Logger - EZTeach Skill
日志工具
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path


class EZTeachLogger:
    """EZTeach日志工具"""
    
    _instance = None
    
    def __new__(cls, log_dir: str = None):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._init_logger(log_dir)
        return cls._instance
    
    def _init_logger(self, log_dir: str = None):
        """初始化日志"""
        self.logger = logging.getLogger('EZTeach')
        self.logger.setLevel(logging.DEBUG)
        
        # 避免重复添加handler
        if self.logger.handlers:
            return
        
        # 控制台输出
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)
        
        # 文件输出
        if log_dir:
            log_path = Path(log_dir)
            log_path.mkdir(parents=True, exist_ok=True)
            log_file = log_path / f"ezteach_{datetime.now().strftime('%Y%m%d')}.log"
            
            file_handler = logging.FileHandler(str(log_file), encoding='utf-8')
            file_handler.setLevel(logging.DEBUG)
            file_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)
    
    def info(self, msg: str):
        self.logger.info(msg)
    
    def debug(self, msg: str):
        self.logger.debug(msg)
    
    def warning(self, msg: str):
        self.logger.warning(msg)
    
    def error(self, msg: str, exc_info: bool = False):
        self.logger.error(msg, exc_info=exc_info)
    
    def success(self, msg: str):
        self.logger.info(f"✓ {msg}")


def get_logger(log_dir: str = None):
    """获取日志实例"""
    return EZTeachLogger(log_dir)
