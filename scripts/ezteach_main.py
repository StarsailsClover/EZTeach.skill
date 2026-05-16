#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EZTeach Main - EZTeach Skill
统一入口脚本，一键调用所有功能，无需人工干预
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import Dict, List, Any, Optional

# 添加模块路径
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

try:
    from ppt_processor import process_ppt_file, PPTProcessor
    from sdm_converter import convert_file, SDMConverter
    from template_adapter import adapt_template, TemplateAdapter
    from doc_generator import generate_document, DocumentGenerator
    ALL_MODULES = True
except ImportError as e:
    ALL_MODULES = False
    IMPORT_ERROR = str(e)


class EZTeachMain:
    """EZTeach主控制器"""
    
    def __init__(self, work_dir: Optional[str] = None):
        self.work_dir = Path(work_dir) if work_dir else Path.cwd()
        self.work_dir.mkdir(parents=True, exist_ok=True)
        
        # 初始化子处理器
        self.sdm_converter = SDMConverter()
        self.doc_generator = DocumentGenerator()
        
        # 工作目录结构
        self.input_dir = self.work_dir / 'input'
        self.output_dir = self.work_dir / 'output'
        self.temp_dir = self.work_dir / 'temp'
        
        for d in [self.input_dir, self.output_dir, self.temp_dir]:
            d.mkdir(parents=True, exist_ok=True)
    
    def process_full_workflow(self, ppt_path: str, template_path: str,
                              lesson_title: str, **kwargs) -> Dict[str, Any]:
        """
        执行完整工作流：PPT提取 → 内容生成 → 模板适配 → 文档输出
        
        Args:
            ppt_path: PPT课件路径
            template_path: 教案模板路径
            lesson_title: 课程标题
            **kwargs: 其他参数
            
        Returns:
            完整处理结果
        """
        result = {
            "success": True,
            "lesson_title": lesson_title,
            "steps": {},
            "output_files": {}
        }
        
        try:
            # 步骤1: 处理PPT
            result["steps"]["ppt_processing"] = {
                "status": "running",
                "input": ppt_path
            }
            
            ppt_processor = PPTProcessor(ppt_path)
            ppt_content = ppt_processor.extract_all_content()
            
            if "error" in ppt_content:
                raise Exception(f"PPT处理失败: {ppt_content['error']}")
            
            result["steps"]["ppt_processing"] = {
                "status": "completed",
                "slides_extracted": ppt_content.get("total_slides", 0),
                "outline": ppt_processor.get_teaching_outline()
            }
            
            # 步骤2: 导出PPT内容为JSON
            content_json = self.temp_dir / f"{lesson_title}_content.json"
            ppt_processor.export_to_json(str(content_json))
            result["output_files"]["content_json"] = str(content_json)
            
            # 步骤3: 解析模板
            result["steps"]["template_analysis"] = {"status": "running"}
            template_adapter = TemplateAdapter(template_path)
            template_structure = template_adapter.parse_template_structure()
            
            if "error" in template_structure:
                raise Exception(f"模板解析失败: {template_structure['error']}")
            
            result["steps"]["template_analysis"] = {
                "status": "completed",
                "sections_found": len(template_structure.get("sections", [])),
                "placeholders_found": len(template_structure.get("placeholders", []))
            }
            
            # 步骤4: 准备内容数据（基于PPT内容）
            content_data = self._prepare_content_data(
                lesson_title, 
                ppt_content, 
                kwargs
            )
            
            # 步骤5: 生成教案
            result["steps"]["document_generation"] = {"status": "running"}
            
            plan_output = self.output_dir / f"{lesson_title}_教案.docx"
            self.doc_generator.generate_teaching_plan(content_data, str(plan_output))
            
            result["steps"]["document_generation"] = {"status": "completed"}
            result["output_files"]["teaching_plan"] = str(plan_output)
            
            # 步骤6: 生成学习单（可选）
            if kwargs.get('generate_sheet', True):
                sheet_output = self.output_dir / f"{lesson_title}_学习单.docx"
                sheet_data = self._prepare_sheet_data(lesson_title, ppt_content)
                self.doc_generator.generate_learning_sheet(sheet_data, str(sheet_output))
                result["output_files"]["learning_sheet"] = str(sheet_output)
            
            # 步骤7: 转换为Markdown备份
            md_output = self.temp_dir / f"{lesson_title}.md"
            self.sdm_converter.docx_to_markdown(str(plan_output), str(md_output))
            result["output_files"]["markdown_backup"] = str(md_output)
            
        except Exception as e:
            result["success"] = False
            result["error"] = str(e)
            for step in result["steps"]:
                if result["steps"][step].get("status") == "running":
                    result["steps"][step]["status"] = "failed"
        
        return result
    
    def _prepare_content_data(self, title: str, ppt_content: Dict, 
                              kwargs: Dict) -> Dict[str, Any]:
        """准备教案内容数据"""
        slides = ppt_content.get("slides", [])
        
        # 提取教学过程
        teaching_process = []
        for slide in slides[1:]:  # 跳过标题页
            slide_title = slide.get("title", "")
            slide_texts = "\n".join(slide.get("texts", []))
            if slide_title:
                teaching_process.append({
                    "title": slide_title,
                    "content": slide_texts,
                    "time": "5分钟"
                })
        
        data = {
            "title": title,
            "basic_info": {
                "subject": kwargs.get('subject', '语文'),
                "grade": kwargs.get('grade', '四年级'),
                "period": kwargs.get('period', '第一课时'),
                "teacher": kwargs.get('teacher', ''),
                "type": kwargs.get('type', '新授课')
            },
            "teaching_objectives": [
                "知识与技能目标：掌握本课核心知识点",
                "过程与方法目标：通过学习培养学生的思维能力",
                "情感态度与价值观：培养学生的学习兴趣"
            ],
            "key_points": "掌握本课的重点内容",
            "difficult_points": "理解本课的难点知识",
            "preparation": "多媒体课件、教案",
            "teaching_process": teaching_process[:8],  # 限制8个步骤
            "blackboard_design": f"{title} - 板书设计",
            "homework": ["完成课后练习", "预习下一课内容"],
            "reflection": "教学反思待填写"
        }
        
        return data
    
    def _prepare_sheet_data(self, title: str, ppt_content: Dict) -> Dict[str, Any]:
        """准备学习单数据"""
        data = {
            "title": f"{title} 学习单",
            "learning_objectives": [
                "能够理解课文主要内容",
                "掌握本课重点字词",
                "能够有感情地朗读课文"
            ],
            "preview": [
                "朗读课文3遍",
                "圈出生字词",
                "思考课文主要讲了什么"
            ],
            "exercises": [
                "抄写本课重点词语：",
                "根据课文内容填空：",
                "回答问题：课文主要讲了什么内容？"
            ],
            "thinking": "学完这篇课文，你有什么收获？"
        }
        return data
    
    def batch_process(self, config_file: str) -> Dict[str, Any]:
        """
        批量处理多个课程
        
        Args:
            config_file: 配置JSON文件路径
            
        Returns:
            批量处理结果
        """
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        results = {
            "total": len(config.get('lessons', [])),
            "success": 0,
            "failed": 0,
            "lessons": []
        }
        
        template_path = config.get('template_path', '')
        
        for lesson in config.get('lessons', []):
            try:
                lesson_result = self.process_full_workflow(
                    ppt_path=lesson['ppt_path'],
                    template_path=template_path,
                    lesson_title=lesson['title'],
                    **lesson.get('options', {})
                )
                if lesson_result['success']:
                    results['success'] += 1
                else:
                    results['failed'] += 1
                results['lessons'].append(lesson_result)
            except Exception as e:
                results['failed'] += 1
                results['lessons'].append({
                    "success": False,
                    "title": lesson.get('title', 'unknown'),
                    "error": str(e)
                })
        
        return results
    
    def convert_format(self, input_path: str, output_path: str, 
                       direction: Optional[str] = None) -> Dict[str, Any]:
        """格式转换"""
        return convert_file(input_path, output_path, direction)
    
    def extract_ppt(self, ppt_path: str, output_json: Optional[str] = None) -> Dict[str, Any]:
        """提取PPT内容"""
        if output_json is None:
            output_json = self.temp_dir / f"{Path(ppt_path).stem}_content.json"
        return process_ppt_file(ppt_path, str(self.temp_dir))
    
    def get_status(self) -> Dict[str, Any]:
        """获取系统状态"""
        return {
            "modules_loaded": ALL_MODULES,
            "import_error": None if ALL_MODULES else IMPORT_ERROR,
            "work_directory": str(self.work_dir),
            "python_version": sys.version,
            "available_commands": [
                "process_full_workflow - 完整工作流",
                "batch_process - 批量处理",
                "convert_format - 格式转换",
                "extract_ppt - PPT内容提取"
            ]
        }


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description='EZTeach - 智能教学文档生成工具')
    parser.add_argument('--mode', choices=['full', 'ppt', 'convert', 'batch', 'status'], 
                       default='status', help='运行模式')
    parser.add_argument('--ppt', help='PPT文件路径')
    parser.add_argument('--template', help='模板文件路径')
    parser.add_argument('--title', help='课程标题')
    parser.add_argument('--input', help='输入文件路径')
    parser.add_argument('--output', help='输出文件路径')
    parser.add_argument('--config', help='批量配置文件')
    parser.add_argument('--work-dir', help='工作目录')
    parser.add_argument('--subject', default='语文', help='学科')
    parser.add_argument('--grade', default='四年级', help='年级')
    
    args = parser.parse_args()
    
    ezteach = EZTeachMain(args.work_dir)
    
    if args.mode == 'status':
        print(json.dumps(ezteach.get_status(), ensure_ascii=False, indent=2))
    
    elif args.mode == 'full':
        if not all([args.ppt, args.template, args.title]):
            print("错误: full模式需要 --ppt, --template, --title 参数")
            sys.exit(1)
        result = ezteach.process_full_workflow(
            args.ppt, args.template, args.title,
            subject=args.subject, grade=args.grade
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))
    
    elif args.mode == 'ppt':
        if not args.ppt:
            print("错误: ppt模式需要 --ppt 参数")
            sys.exit(1)
        result = ezteach.extract_ppt(args.ppt, args.output)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    
    elif args.mode == 'convert':
        if not all([args.input, args.output]):
            print("错误: convert模式需要 --input 和 --output 参数")
            sys.exit(1)
        result = ezteach.convert_format(args.input, args.output)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    
    elif args.mode == 'batch':
        if not args.config:
            print("错误: batch模式需要 --config 参数")
            sys.exit(1)
        result = ezteach.batch_process(args.config)
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
