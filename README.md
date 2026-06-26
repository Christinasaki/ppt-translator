# PPT一键翻译 (PPT Translator)

通用的 PowerPoint 演示文稿翻译工具，支持中英双向翻译，可配置术语库。

## 功能特性

- **双向翻译**：支持中文→英文、英文→中文
- **可配置术语库**：通过 Excel 或 JSON 文件提供专业术语映射
- **智能排版**：自动调整字体大小，防止文本溢出
- **SmartArt 支持**：翻译 SmartArt 图形中的文本
- **组合图形支持**：正确处理嵌套组合图形的坐标变换
- **跨平台**：支持 Windows、macOS、Linux
- **多种使用方式**：
  - Python 包（pip 安装）
  - 命令行工具
  - 独立脚本（单文件版本）
  - QoderWork Skill

## 安装

### 方式一：pip 安装

```bash
pip install ppt-translator
```

### 方式二：从源码安装

```bash
git clone https://github.com/Christinasaki/ppt-translator.git
cd ppt-translator
pip install -e .
```

### 方式三：独立脚本

下载 `ppt_translate_standalone.py`，直接运行（需先安装依赖）：

```bash
pip install python-pptx Pillow openpyxl lxml
python ppt_translate_standalone.py input.pptx -o output.pptx
```

## 使用方法

### 命令行

```bash
# 基本用法：中文→英文
ppt-translate input.pptx -o output.pptx

# 指定术语库
ppt-translate input.pptx --terms terms.xlsx -o output.pptx

# 英文→中文
ppt-translate input_en.pptx --direction en2zh --terms terms.xlsx -o output_zh.pptx

# 使用配置文件
ppt-translate input.pptx --config config.json

# 禁用 Title Case
ppt-translate input.pptx --no-title-case
```

### Python API

```python
from ppt_translator import PPTTranslator, Settings

# 创建配置
settings = Settings(
    direction="zh2en",
    target_font="Arial",
    terminology_file="terms.xlsx",
)

# 创建翻译器并执行
translator = PPTTranslator(settings)
translator.translate_ppt("input.pptx", "output.pptx")
```

## 术语库格式

### Excel 格式

| 中文 | 英文 |
|------|------|
| 人工智能 | Artificial Intelligence |
| 机器学习 | Machine Learning |
| 深度学习 | Deep Learning |

### JSON 格式

```json
{
  "人工智能": "Artificial Intelligence",
  "机器学习": "Machine Learning",
  "深度学习": "Deep Learning"
}
```

## 配置文件

创建 `config.json`：

```json
{
  "direction": "zh2en",
  "source_font": "Arial",
  "target_font": "Arial",
  "terminology_file": "terms.xlsx",
  "short_translation_threshold": {
    "width_inches": 5.0,
    "shallow_width": 6.0,
    "shallow_height": 1.0
  },
  "font_fitting": {
    "margin_emu": 60000,
    "line_height_ratio": 1.35,
    "min_font_size": 5
  },
  "title_case": {
    "enabled": true,
    "small_words": ["a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet", "at", "by", "from", "in", "into", "of", "off", "on", "onto", "per", "to", "up", "via", "with", "as"]
  }
}
```

## 技术特性

### 字体处理

- 统一设置 Latin、EastAsian、ComplexScript 三种字体
- 自动检测系统字体路径
- 防止 PowerPoint 回退到主题字体

### 溢出防护

- 扣除文本框边距（默认 60000 EMU）
- 使用 1.35 倍行高估算（更接近 PowerPoint 实际渲染）
- 自适应缩小字号（最小 5pt）
- 根据文本框尺寸自动选择简短/完整翻译

### 坐标变换

- 正确处理组合图形的嵌套坐标系统
- 累积计算 `a:ext / a:chExt` 缩放因子
- 确保嵌套形状的尺寸计算准确

## 依赖

- Python >= 3.8
- python-pptx >= 0.6.21
- Pillow >= 9.0.0
- openpyxl >= 3.0.0
- lxml >= 4.9.0

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 致谢

本项目基于实际 PPT 翻译需求开发，感谢所有反馈问题的用户。
