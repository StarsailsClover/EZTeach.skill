# EZTeach Skill 配置说明

## 环境要求

### 系统要求
- **操作系统**: Windows 10+ / macOS 10.14+ / Linux (Ubuntu 18.04+)
- **Python版本**: 3.8 - 3.12
- **内存**: 建议 4GB 以上
- **磁盘**: 至少 500MB 可用空间

### Python依赖

```
python-docx >= 0.8.11      # Word文档处理
python-pptx >= 0.6.21      # PPT文档处理
markdown >= 3.4            # Markdown处理
pywin32 >= 305             # Windows COM接口 (仅Windows)
```

## 安装部署

### 方式一：标准安装

```bash
# 1. 解压技能包到指定目录
unzip EZTeach.skill.zip -d /path/to/skills/

# 2. 进入目录
cd /path/to/skills/EZTeach.skill

# 3. 安装依赖
pip install -r requirements.txt

# 4. 验证安装
python scripts/ezteach_main.py --mode status
```

### 方式二：虚拟环境安装（推荐）

```bash
# 1. 创建虚拟环境
python -m venv venv

# 2. 激活虚拟环境
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

# 3. 安装依赖
pip install -r requirements.txt
```

### 验证安装成功

运行状态检查命令：

```bash
python scripts/ezteach_main.py --mode status
```

预期输出：
```json
{
  "modules_loaded": true,
  "work_directory": "...",
  "available_commands": [...]
}
```

## 依赖安装详解

### 核心依赖安装

```bash
# 文档处理核心库
pip install python-docx python-pptx

# Markdown处理
pip install markdown
```

### Windows PDF转换支持

```bash
# 安装pywin32用于PPT转PDF
pip install pywin32

# 验证PowerPoint
# 需要安装 Microsoft PowerPoint 2016+
```

### Linux PDF转换支持

```bash
# 使用LibreOffice
sudo apt install libreoffice

# 转换命令示例
libreoffice --headless --convert-to pdf input.pptx
```

## 目录配置

### 工作目录结构

运行时自动创建以下目录：

```
工作目录/
├── input/      # 输入文件（PPT、模板）
├── output/     # 输出文件（教案、学习单）
├── temp/       # 临时文件（JSON、MD备份）
└── logs/       # 日志文件（可选）
```

### 自定义工作目录

通过命令行参数指定：

```bash
python scripts/ezteach_main.py --mode full \
  --work-dir /custom/path \
  --ppt ...
```

## 配置文件

### 批量处理配置 (batch_config.json)

```json
{
  "template_path": "./templates/standard_lesson_plan.docx",
  "default_options": {
    "subject": "语文",
    "grade": "四年级",
    "generate_sheet": true
  },
  "lessons": [
    {
      "title": "繁星 第一课时",
      "ppt_path": "./input/繁星.pptx",
      "options": {
        "period": "第一课时"
      }
    }
  ]
}
```

### 样式配置

修改 `doc_generator.py` 中的 `default_styles` 字典自定义样式：

```python
self.default_styles = {
    'title': {
        'font_size': 18,
        'bold': True,
        'alignment': WD_ALIGN_PARAGRAPH.CENTER,
        'space_after': 24
    }
}
```

## 环境变量配置

| 环境变量 | 说明 | 默认值 |
|----------|------|--------|
| `EZTEACH_WORK_DIR` | 默认工作目录 | 当前目录 |
| `EZTEACH_LOG_LEVEL` | 日志级别 (DEBUG/INFO/WARN/ERROR) | INFO |
| `EZTEACH_TEMP_DIR` | 临时文件目录 | ./temp |

## 常见问题排查

### 问题1：ImportError: No module named 'docx'

**原因**: python-docx 未安装

**解决**:
```bash
pip install python-docx
```

### 问题2：PPT提取内容为空

**原因**: 
1. PPT文件格式为旧版 .ppt
2. PPT内容为图片形式（无文本）

**解决**:
1. 将 .ppt 另存为 .pptx 格式
2. 确保PPT包含可编辑文本

### 问题3：生成的文档乱码

**原因**: 编码问题

**解决**:
- 确保所有源文件使用 UTF-8 编码
- Windows系统确保区域设置支持中文

### 问题4：PDF转换失败

**原因**: 
- Windows: 未安装PowerPoint或pywin32
- Linux: 未安装LibreOffice

**解决**:
```bash
# Windows
pip install pywin32

# Linux
sudo apt install libreoffice
```

### 问题5：模板占位符不替换

**原因**: 占位符格式不正确

**解决**:
- 确保占位符格式为 `{{PLACEHOLDER_NAME}}`
- 占位符不要跨多个Word Run
- 建议先清除格式再输入占位符

## 性能优化

### 大文件处理

对于超过50页的PPT：

```bash
# 增加Python堆内存
python -X pycache_prefix=./cache scripts/ezteach_main.py ...
```

### 批量处理优化

```python
# 在config中启用并行处理（未来版本）
"batch_options": {
    "parallel": true,
    "max_workers": 4
}
```

### 临时文件清理

```bash
# 自动清理超过1小时的临时文件
python -c "from libs.file_utils import FileUtils; FileUtils.cleanup_dir('./temp', older_than=3600)"
```

## 卸载

```bash
# 1. 删除依赖
pip uninstall -y python-docx python-pptx markdown pywin32

# 2. 删除技能目录
rm -rf /path/to/EZTeach.skill
```

## 更新升级

```bash
# 1. 备份配置
cp -r EZTeach.skill EZTeach.skill.backup

# 2. 解压新版本
unzip EZTeach.skill.new.zip

# 3. 迁移配置
# 复制自定义模板和配置文件
```
