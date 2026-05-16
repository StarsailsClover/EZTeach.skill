#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Docx Utils - EZTeach Skill
Word文档工具类
"""

import os
from typing import Dict, List, Any, Optional
from pathlib import Path
from copy import deepcopy


class DocxUtils:
    """Word文档工具类"""
    
    @staticmethod
    def get_text(doc) -> str:
        """提取文档所有文本"""
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    
    @staticmethod
    def get_styles(doc) -> Dict[str, Any]:
        """获取文档所有样式"""
        styles = {}
        for style in doc.styles:
            styles[style.name] = {
                'type': style.type,
                'builtin': style.builtin
            }
        return styles
    
    @staticmethod
    def clone_paragraph_style(source_para, target_para) -> None:
        """克隆段落样式"""
        target_para.style = source_para.style
        target_para.alignment = source_para.alignment
        target_para.paragraph_format.left_indent = source_para.paragraph_format.left_indent
        target_para.paragraph_format.right_indent = source_para.paragraph_format.right_indent
        target_para.paragraph_format.first_line_indent = source_para.paragraph_format.first_line_indent
        target_para.paragraph_format.space_before = source_para.paragraph_format.space_before
        target_para.paragraph_format.space_after = source_para.paragraph_format.space_after
        target_para.paragraph_format.line_spacing = source_para.paragraph_format.line_spacing
    
    @staticmethod
    def clone_run_format(source_run, target_run) -> None:
        """克隆Run格式"""
        target_run.font.name = source_run.font.name
        target_run.font.size = source_run.font.size
        target_run.font.bold = source_run.font.bold
        target_run.font.italic = source_run.font.italic
        target_run.font.underline = source_run.font.underline
        target_run.font.color.rgb = source_run.font.color.rgb if source_run.font.color else None
    
    @staticmethod
    def clear_document(doc) -> None:
        """清空文档内容（保留样式）"""
        for para in doc.paragraphs:
            para.clear()
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        para.clear()
    
    @staticmethod
    def find_placeholder(doc, pattern: str = r'\{\{([^}]+)\}\}') -> List[Dict]:
        """查找文档中的占位符"""
        import re
        pattern = re.compile(pattern)
        placeholders = []
        
        for i, para in enumerate(doc.paragraphs):
            matches = pattern.findall(para.text)
            for match in matches:
                placeholders.append({
                    'name': match,
                    'paragraph_index': i,
                    'paragraph_text': para.text,
                    'style': para.style.name
                })
        
        # 查找表格中的占位符
        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    for para in cell.paragraphs:
                        matches = pattern.findall(para.text)
                        for match in matches:
                            placeholders.append({
                                'name': match,
                                'table_index': table_idx,
                                'row': row_idx,
                                'col': col_idx,
                                'cell_text': cell.text
                            })
        
        return placeholders
    
    @staticmethod
    def replace_text(doc, old_text: str, new_text: str) -> int:
        """替换文档中的文本"""
        count = 0
        
        # 替换段落中的文本
        for para in doc.paragraphs:
            if old_text in para.text:
                for run in para.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                        count += 1
        
        # 替换表格中的文本
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if old_text in para.text:
                            for run in para.runs:
                                if old_text in run.text:
                                    run.text = run.text.replace(old_text, new_text)
                                    count += 1
        
        return count
