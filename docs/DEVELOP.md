# EZTeach Skill 开发文档

## 模块架构

### 整体架构

```
EZTeach Skill
├── 入口层 (ezteach_main.py)
│   └── EZTeachMain 主控制器
│       ├── 工作流调度
│       ├── 参数解析
│       └── 状态管理
├── 核心层 (scripts/)
│   ├── PPTProcessor      # PPT处理
│   ├── SDMConverter      # 格式转换
│   ├── TemplateAdapter   # 模板适配
│   └── DocumentGenerator # 文档生成
└── 工具层 (libs/)
    ├── DocxUtils         # Word工具
    ├── PPTUtils          # PPT工具
    ├── FileUtils         # 文件工具
    └── Logger            # 日志工具
```

## 核心模块说明

### 1. PPT处理器 (ppt_processor.py)

**类：PPTProcessor**

负责PPT内容的完整提取和分析。

#### 主要方法：

```python
# 提取所有幻灯片内容
extract_all_content() -> Dict

# 解析动画窗格信息
parse_animations() -> List[Dict]

# 导出为JSON
export_to_json(output_path: str) -> str

# 转换为PDF
convert_to_pdf(output_path: str) -> str

# 提取教学大纲
get_teaching_outline() -> List[str]
```

#### 输出数据结构：

```python
{
  "file_name": "课件.pptx",
  "total_slides": 28,
  "slides": [
    {
      "slide_number": 1,
      "title": "课程标题",
      "texts": ["文本1", "文本2"],
      "images": [],
      "tables": [],
      "notes": "备注内容"
    }
  ],
  "animations": [...]
}
```

### 2. SDM转换器 (sdm_converter.py)

**类：SDMConverter**

实现Word和Markdown之间的双向无损转换。

#### 主要方法：

```python
# Word → Markdown
docx_to_markdown(docx_path: str, output_path: str) -> str

# Markdown → Word
markdown_to_docx(md_path: str, output_path: str, template_path: str) -> str

# 捕获格式语义
capture_format_semantics(docx_path: str) -> Dict

# 批量转换
batch_convert(input_dir: str, output_dir: str, direction: str) -> Dict
```

#### 转换特性：
- ✅ 保留标题层级
- ✅ 保留列表格式
- ✅ 保留粗体/斜体
- ✅ 保留表格结构
- ✅ 捕获样式元数据

### 3. 模板适配器 (template_adapter.py)

**类：TemplateAdapter**

解析模板结构，实现内容的智能映射。

#### 主要方法：

```python
# 解析模板结构
parse_template_structure(template_path: str) -> Dict

# 提取占位符Schema
extract_placeholder_schema() -> Dict

# 内容映射到模板
map_content_to_template(content_data: Dict, output_path: str) -> str

# 克隆格式
clone_format_from_template(source: str, target: str, output: str) -> str
```

#### 占位符识别：
- 自动识别 `{{PLACEHOLDER}}` 格式
- 支持段落和表格内占位符
- 替换时保留原有样式

### 4. 文档生成器 (doc_generator.py)

**类：DocumentGenerator**

完全脚本控制的文档生成，不依赖AI处理格式。

#### 主要方法：

```python
# 生成教案
generate_teaching_plan(plan_data: Dict, output_path: str) -> str

# 生成学习单
generate_learning_sheet(sheet_data: Dict, output_path: str) -> str

# 批量生成
batch_generate(data_list: List, output_dir: str, doc_type: str) -> Dict
```

#### 内置样式系统：
- title: 18号，加粗，居中
- heading1/2/3: 分级标题
- normal: 12号，1.5倍行距，首行缩进
- list: 左缩进，项目符号
- table_header/cell: 表格样式

### 5. 主控制器 (ezteach_main.py)

**类：EZTeachMain**

统一入口，一键调用所有功能。

#### 工作流：

```
1. PPT提取 → 2. 模板解析 → 3. 内容准备 → 4. 文档生成 → 5. 格式转换
```

#### 主要方法：

```python
# 完整工作流
process_full_workflow(ppt_path, template_path, title, **kwargs) -> Dict

# 批量处理
batch_process(config_file: str) -> Dict

# 格式转换
convert_format(input_path, output_path, direction) -> Dict

# 状态检查
get_status() -> Dict
```

## 扩展开发指南

### 添加新的文档类型

1. 在 `doc_generator.py` 中添加新的生成方法：

```python
def generate_exam_paper(self, data: Dict, output_path: str) -> str:
    self.create_new_document()
    # 添加你的生成逻辑...
    self.doc.save(output_path)
    return output_path
```

2. 在 `ezteach_main.py` 中添加调用入口

3. 更新 `SKILL.md` 使用说明

### 添加新的占位符规则

1. 修改 `template_adapter.py` 中的 `placeholder_pattern`：

```python
self.placeholder_pattern = re.compile(r'\$\{([^}]+)\}')  # 支持${VAR}格式
```

2. 更新占位符检测和替换逻辑

### 自定义样式

1. 修改 `doc_generator.py` 中的 `default_styles` 字典：

```python
self.default_styles = {
    'custom_style': {
        'font_size': 14,
        'bold': True,
        'alignment': WD_ALIGN_PARAGRAPH.RIGHT
    }
}
```

2. 添加对应的 `_apply_style` 分支

## 代码规范

### Python代码规范
- 遵循 PEP 8 规范
- 使用类型注解
- 每个函数必须有docstring
- 异常处理要完善

### 文件命名规范
- 模块名：小写+下划线（snake_case）
- 类名：大驼峰（PascalCase）
- 函数名：小写+下划线（snake_case）

### 注释规范
- 每个模块头部必须有功能说明
- 每个公共方法必须有参数和返回值说明
- 复杂逻辑必须添加行内注释

## 调试技巧

### 1. 启用详细日志

```python
from libs.logger import get_logger
logger = get_logger('./logs')
logger.debug('调试信息')
```

### 2. 单独测试模块

```bash
# 单独测试PPT处理
python -c "from scripts.ppt_processor import PPTProcessor; p = PPTProcessor('test.pptx'); print(p.extract_all_content())"
```

### 3. 检查中间输出

运行后检查 `temp/` 目录下的JSON文件，验证数据提取是否正确。

## 版本历史

- **v2.0.0** (当前)
  - 重构完整目录结构
  - 所有功能脚本化
  - 新增主控制器
  - 完善工具类库

- **v1.0.0**
  - 初始版本
  - 基础PPT提取
  - 简单文档生成
