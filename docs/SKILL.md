# EZTeach Skill 使用说明

## 概述

EZTeach 是一个智能教学文档生成AI Skill，能够自动解析PPT课件内容，基于教案模板一键生成完整的教案和学习单文档。

## 功能特性

- ✅ **PPT智能解析**：自动提取PPT内容、大纲、备注、动画信息
- ✅ **格式双向转换**：Word ↔ Markdown 无损转换，保持格式语义
- ✅ **模板智能适配**：自动解析模板结构，100%复用格式样式
- ✅ **文档自动生成**：脚本控制格式，生成标准教案和学习单
- ✅ **一键工作流**：完整自动化流程，无需人工干预

## 目录结构

```
EZTeach.skill/
├── scripts/              # 核心Python脚本
│   ├── ezteach_main.py   # 主入口，一键调用
│   ├── ppt_processor.py  # PPT处理模块
│   ├── sdm_converter.py  # 格式转换模块
│   ├── template_adapter.py # 模板适配模块
│   └── doc_generator.py  # 文档生成模块
├── libs/                 # 工具类库
│   ├── docx_utils.py     # Word文档工具
│   ├── ppt_utils.py      # PPT处理工具
│   ├── file_utils.py     # 文件操作工具
│   └── logger.py         # 日志工具
├── prompts/              # 系统提示词
│   ├── system_prompt.txt
│   ├── content_generation.txt
│   └── template_mapping.txt
├── docs/                 # 文档
│   ├── SKILL.md          # 使用说明（本文件）
│   ├── DEVELOP.md        # 开发文档
│   └── CONFIG.md         # 配置说明
├── skill.json            # Skill入口配置
├── requirements.txt      # Python依赖
└── README.md             # 项目说明
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 完整工作流（一键生成教案）

```bash
python scripts/ezteach_main.py --mode full \
  --ppt ./课件/繁星.pptx \
  --template ./模板/教案模板.docx \
  --title "繁星 第一课时" \
  --subject "语文" \
  --grade "四年级"
```

### 3. 单独功能调用

#### 仅提取PPT内容
```bash
python scripts/ezteach_main.py --mode ppt --ppt ./课件.pptx
```

#### 格式转换
```bash
# Word转Markdown
python scripts/ezteach_main.py --mode convert --input input.docx --output output.md

# Markdown转Word
python scripts/ezteach_main.py --mode convert --input input.md --output output.docx
```

#### 批量处理
```bash
python scripts/ezteach_main.py --mode batch --config batch_config.json
```

## 参数说明

### 主入口参数 (ezteach_main.py)

| 参数 | 说明 | 必填 | 默认值 |
|------|------|------|--------|
| `--mode` | 运行模式：full/ppt/convert/batch/status | 否 | status |
| `--ppt` | PPT课件文件路径 | full/ppt模式必填 | - |
| `--template` | 教案模板文件路径 | full模式必填 | - |
| `--title` | 课程标题 | full模式必填 | - |
| `--input` | 输入文件路径 | convert模式必填 | - |
| `--output` | 输出文件路径 | convert模式可选 | - |
| `--config` | 批量配置JSON文件 | batch模式必填 | - |
| `--work-dir` | 工作目录 | 否 | 当前目录 |
| `--subject` | 学科 | 否 | 语文 |
| `--grade` | 年级 | 否 | 四年级 |

### 批量配置文件格式 (batch_config.json)

```json
{
  "template_path": "./模板/通用教案模板.docx",
  "lessons": [
    {
      "title": "课程1",
      "ppt_path": "./课件/课程1.pptx",
      "options": {
        "subject": "语文",
        "grade": "四年级"
      }
    }
  ]
}
```

## 输出文件

运行后将在工作目录生成：

```
output/
├── {课程名}_教案.docx      # 完整教案文档
└── {课程名}_学习单.docx    # 配套学习单

temp/
├── {课程名}_content.json   # PPT提取内容
└── {课程名}.md            # Markdown备份
```

## 使用示例

### 示例1：生成语文教案

```bash
python scripts/ezteach_main.py --mode full \
  --ppt "./小学语文/繁星/课件.pptx" \
  --template "./模板/小学语文教案模板.docx" \
  --title "繁星 第一课时" \
  --subject "语文" \
  --grade "四年级"
```

### 示例2：生成数学学习单

```bash
python scripts/doc_generator.py \
  ./data/三角形认识.json \
  ./output/三角形认识_学习单.docx \
  sheet
```

## 注意事项

1. **脚本优先原则**：所有文件操作、格式处理均由Python脚本完成，AI模型仅负责语义内容生成
2. **无测试文件**：Skill包内不包含任何测试输出文件，所有输出在运行时生成
3. **环境要求**：PDF转换功能需要Windows环境 + pywin32，其他平台可使用LibreOffice
4. **格式兼容**：支持 .pptx / .docx / .md 格式，不支持旧版 .ppt / .doc 格式

## 常见问题

**Q: 提示python-docx未安装？**
A: 执行 `pip install python-docx python-pptx` 安装依赖

**Q: PPT内容提取不完整？**
A: 确保PPT是.pptx格式，旧版.ppt需要先转换格式

**Q: 生成的文档格式错乱？**
A: 确保模板文件格式规范，使用标准Word样式

**Q: 如何自定义教案结构？**
A: 修改模板文件的结构，模板适配器会自动解析适配
