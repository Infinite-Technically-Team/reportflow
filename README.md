# Excel报表生成器

一个快速生成带图表的Excel报告和HTML报告的Python工具，输入数据→3秒生成带图表报告 + index.html

## 功能特性

- 🚀 **快速生成**：3秒内完成报告生成
- 📊 **丰富图表**：支持柱状图、折线图、直方图、散点图、饼图等多种图表
- 📈 **Excel报告**：生成包含数据表格、统计摘要和图表的Excel文件
- 🌐 **HTML报告**：生成美观的响应式HTML报告，支持直接在浏览器中查看
- 📁 **多种输入格式**：支持CSV、Excel、JSON文件，以及字典、列表、DataFrame等数据结构
- 💻 **命令行 + 交互式**：支持命令行参数和交互式两种使用方式

## 安装依赖

```bash
pip install -r requirements.txt
```

## 快速开始

### 1. 使用命令行

#### 使用示例数据

```bash
python report_generator.py --demo
```

#### 从文件加载数据

```bash
python report_generator.py -i sample_data.csv -e my_report.xlsx -t my_index.html
```

### 2. 使用交互式模式

```bash
python report_generator.py
```

然后按照提示选择输入方式。

### 3. 在Python代码中使用

```python
from report_generator import ReportGenerator

# 创建生成器实例
generator = ReportGenerator()

# 使用字典数据
data = {
    "月份": ["1月", "2月", "3月"],
    "销售额": [12000, 15000, 18000],
    "利润": [3600, 4500, 5400]
}

# 生成报告
generator.generate_report(data)
```

## 项目结构

```
.
├── report_generator.py    # 主程序入口
├── data_processor.py      # 数据处理模块
├── excel_generator.py     # Excel生成模块
├── html_generator.py      # HTML生成模块
├── sample_data.csv        # 示例数据
├── requirements.txt       # 依赖文件
└── README.md              # 项目说明
```

## 生成的报告内容

### Excel报告

- **数据**：完整的数据表格，带格式化表头和自动调整的列宽
- **统计摘要**：包含均值、标准差、最大值、最小值等统计信息
- **图表**：自动生成的柱状图和折线图

### HTML报告

- **数据概览**：显示数据行数、列数和关键统计指标
- **图表分析**：多种图表类型，支持响应式显示
- **数据表格**：显示前10行数据，支持滚动查看
- **统计摘要**：完整的统计信息表格

## 性能优化

为了确保3秒内完成生成，建议：

- 数据量控制在10万行以内
- 数值列数量控制在10列以内
- 图表数量控制在5个以内

## 支持的数据格式

### 输入文件格式

- CSV文件 (.csv)
- Excel文件 (.xlsx, .xls)
- JSON文件 (.json)

### JSON数据格式

支持两种JSON格式：

1. 数组格式：
```json
[
    {"月份": "1月", "销售额": 12000},
    {"月份": "2月", "销售额": 15000}
]
```

2. 字典格式：
```json
{
    "columns": ["月份", "销售额"],
    "data": [
        ["1月", 12000],
        ["2月", 15000]
    ]
}
```

或直接的字典格式：
```json
{
    "月份": ["1月", "2月"],
    "销售额": [12000, 15000]
}
```

## 自定义输出

### 自定义Excel输出路径

```bash
python report_generator.py -i sample_data.csv -e custom_report.xlsx
```

### 自定义HTML输出路径

```bash
python report_generator.py -i sample_data.csv -t custom_index.html
```

## 技术栈

- **Python 3.7+**：主要开发语言
- **pandas**：数据处理和分析
- **openpyxl**：Excel文件生成
- **matplotlib**：图表生成
- **jinja2**：HTML模板渲染
- **base64**：图表编码

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request！

## 更新日志

### v1.0.0

- 初始版本
- 支持Excel和HTML报告生成
- 支持多种图表类型
- 支持多种数据输入格式
- 支持命令行和交互式使用
