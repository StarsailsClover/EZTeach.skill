#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT Utils - EZTeach Skill
PPT处理工具类
"""

import os
from typing import Dict, List, Any, Optional
from pathlib import Path


class PPTUtils:
    """PPT处理工具类"""
    
    @staticmethod
    def extract_text_from_shape(shape) -> str:
        """从形状中提取文本"""
        text = []
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text.append(paragraph.text)
        return '\n'.join(text)
    
    @staticmethod
    def get_slide_titles(prs) -> List[str]:
        """获取所有幻灯片标题"""
        titles = []
        for slide in prs.slides:
            if slide.shapes.title:
                titles.append(slide.shapes.title.text)
            else:
                titles.append('')
        return titles
    
    @staticmethod
    def count_slide_elements(slide) -> Dict[str, int]:
        """统计幻灯片元素数量"""
        counts = {
            'shapes': 0,
            'text_frames': 0,
            'pictures': 0,
            'tables': 0,
            'charts': 0,
            'groups': 0
        }
        
        for shape in slide.shapes:
            counts['shapes'] += 1
            if shape.has_text_frame:
                counts['text_frames'] += 1
            if shape.shape_type == 13:  # PICTURE
                counts['pictures'] += 1
            if shape.has_table:
                counts['tables'] += 1
            if shape.has_chart:
                counts['charts'] += 1
            if shape.shape_type == 6:  # GROUP
                counts['groups'] += 1
        
        return counts
    
    @staticmethod
    def extract_all_text(prs) -> str:
        """提取PPT所有文本"""
        all_text = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.text.strip():
                            all_text.append(paragraph.text.strip())
        return '\n'.join(all_text)
    
    @staticmethod
    def extract_notes(prs) -> Dict[int, str]:
        """提取所有备注"""
        notes = {}
        for i, slide in enumerate(prs.slides):
            if slide.has_notes_slide:
                notes_slide = slide.notes_slide
                if notes_slide.notes_text_frame:
                    notes[i + 1] = notes_slide.notes_text_frame.text
        return notes
    
    @staticmethod
    def get_slide_size(prs) -> Dict[str, float]:
        """获取幻灯片尺寸"""
        return {
            'width': prs.slide_width,
            'height': prs.slide_height,
            'aspect_ratio': prs.slide_width / prs.slide_height if prs.slide_height else 0
        }
