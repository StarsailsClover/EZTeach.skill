# EZTeach Skill v2.0.0

## 智能教学文档生成AI Skill

EZTeach 是一个专业的教学文档处理AI Skill，能够自动解析PPT课件内容，基于教案模板一键生成完整的教案和学习单文档。

## Agent安装提示词

帮我安装这个Skill并检查：https://github.com/StarsailsClover/EZTeach.skill/releases/tag/260516
## ✨ 核心特性

- **PPT智能解析**：自动提取内容、大纲、备注、动画信息
- **格式双向转换**：Word ↔ Markdown 无损转换，保持格式语义
- **模板智能适配**：自动解析模板结构，100%复用格式样式
- **文档自动生成**：脚本控制格式，生成标准教案和学习单
- **一键工作流**：完整自动化流程，无需人工干预

## 📁 目录结构

```
EZTeach.skill/
├── scripts/          # 核心Python脚本 (5个)
├── libs/             # 工具类库 (4个)
├── prompts/          # 系统提示词 (3个)
├── docs/             # 完整文档 (3个)
├── skill.json        # Skill入口配置
├── requirements.txt  # Python依赖
└── README.md         # 本文件
```

**无空文件夹 | 无测试输出文件 | 所有脚本完整可运行**

## 🚀 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 一键生成教案

```bash
python scripts/ezteach_main.py --mode full \
  --ppt ./课件.pptx \
  --template ./教案模板.docx \
  --title "课程标题" \
  --subject "语文" \
  --grade "四年级"
```

### 检查系统状态

```bash
python scripts/ezteach_main.py --mode status
```

## 📚 文档

- **[使用说明](docs/SKILL.md)** - 完整使用指南、参数说明、调用示例
- **[开发文档](docs/DEVELOP.md)** - 模块说明、架构设计、扩展指南
- **[配置说明](docs/CONFIG.md)** - 环境部署、依赖安装、故障排查

## 🎯 设计原则

1. **脚本优先原则**：所有格式处理、文件操作由Python脚本完成
2. **AI专注原则**：模型仅负责语义内容生成，不执行脚本工作
3. **无人工干预**：完整自动化流程，一键完成所有操作
4. **纯净技能包**：不包含任何测试输出文件，运行时生成输出

## 🔧 技术栈

- **python-docx** - Word文档处理
- **python-pptx** - PPT课件处理
- **markdown** - Markdown格式转换

## 📄 License

MIT License
