import pandas as pd
import matplotlib.pyplot as plt
import base64
import io
import time
from jinja2 import Template

class HTMLGenerator:
    def __init__(self):
        self.data = None
        self.summary_data = None
        self.charts = []
    
    def set_data(self, data: pd.DataFrame) -> None:
        """
        设置数据
        :param data: 要生成报告的数据
        """
        self.data = data
        self.summary_data = data.describe()
    
    def generate_chart_base64(self, chart_type: str, title: str, x_column: str = None, 
                             y_columns: list = None, figsize: tuple = (12, 6)) -> str:
        """
        生成Base64编码的图表
        :param chart_type: 图表类型
        :param title: 图表标题
        :param x_column: X轴列名
        :param y_columns: Y轴列名列表
        :param figsize: 图表大小
        :return: Base64编码的图表
        """
        if self.data is None:
            raise ValueError("没有设置数据，请先调用set_data方法")
        
        # 如果没有指定列，使用默认列
        if x_column is None:
            x_column = self.data.columns[0]
        if y_columns is None:
            y_columns = self.data.columns[1:3] if len(self.data.columns) > 1 else [self.data.columns[0]]
        
        # 创建图表
        plt.figure(figsize=figsize, dpi=100)
        plt.title(title, fontsize=16, fontweight='bold')
        
        if chart_type == "bar":
            self.data.plot(x=x_column, y=y_columns, kind="bar", ax=plt.gca())
        elif chart_type == "line":
            self.data.plot(x=x_column, y=y_columns, kind="line", ax=plt.gca(), marker='o')
        elif chart_type == "scatter":
            if len(y_columns) > 0:
                for y_col in y_columns:
                    plt.scatter(self.data[x_column], self.data[y_col], label=y_col, alpha=0.7)
                plt.legend()
        elif chart_type == "histogram":
            if len(y_columns) > 0:
                for y_col in y_columns:
                    plt.hist(self.data[y_col].dropna(), alpha=0.5, label=y_col, bins=20)
                plt.legend()
        elif chart_type == "pie":
            if len(y_columns) == 1:
                pie_data = self.data[y_columns[0]].value_counts()
                plt.pie(pie_data.values, labels=pie_data.index, autopct='%1.1f%%', startangle=90)
                plt.axis('equal')
        else:
            raise ValueError("不支持的图表类型")
        
        plt.xlabel(x_column, fontsize=12)
        plt.ylabel("值", fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        # 保存图表到内存并编码为Base64
        img_stream = io.BytesIO()
        plt.savefig(img_stream, format='png', bbox_inches='tight')
        img_stream.seek(0)
        img_base64 = base64.b64encode(img_stream.read()).decode('utf-8')
        plt.close()
        
        return f"data:image/png;base64,{img_base64}"
    
    def generate_html(self, output_path: str = "index.html") -> str:
        """
        生成完整的HTML报告
        :param output_path: 输出路径
        :return: 生成的文件路径
        """
        if self.data is None:
            raise ValueError("没有设置数据，请先调用set_data方法")
        
        start_time = time.time()
        
        # 生成图表
        numeric_cols = self.data.select_dtypes(include=['number']).columns.tolist()
        charts = []
        
        if len(numeric_cols) > 0:
            # 柱状图
            bar_chart = self.generate_chart_base64("bar", "柱状图分析", 
                                                  x_column=self.data.columns[0], 
                                                  y_columns=numeric_cols[:3])
            charts.append({"type": "bar", "title": "柱状图分析", "image": bar_chart})
            
            # 折线图
            line_chart = self.generate_chart_base64("line", "折线图分析", 
                                                  x_column=self.data.columns[0], 
                                                  y_columns=numeric_cols[:3])
            charts.append({"type": "line", "title": "折线图分析", "image": line_chart})
            
            # 直方图
            if len(numeric_cols) > 0:
                hist_chart = self.generate_chart_base64("histogram", "数据分布直方图", 
                                                       y_columns=numeric_cols[:3])
                charts.append({"type": "histogram", "title": "数据分布直方图", "image": hist_chart})
        
        # HTML模板
        html_template = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>数据报告</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            background-color: #f5f7fa;
            color: #333;
            line-height: 1.6;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
        }
        
        header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px 0;
            text-align: center;
            border-radius: 10px;
            margin-bottom: 30px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        
        h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .section {
            background: white;
            padding: 30px;
            margin-bottom: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        }
        
        h2 {
            color: #667eea;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 2px solid #e0e0e0;
            padding-bottom: 10px;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        
        .stat-card {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        
        .stat-card h3 {
            font-size: 1.2em;
            margin-bottom: 10px;
            opacity: 0.9;
        }
        
        .stat-card .value {
            font-size: 2em;
            font-weight: 700;
        }
        
        .chart-container {
            margin-bottom: 30px;
            text-align: center;
        }
        
        .chart-container h3 {
            color: #333;
            margin-bottom: 15px;
            font-size: 1.4em;
        }
        
        .chart-container img {
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 0.9em;
        }
        
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e0e0e0;
        }
        
        th {
            background-color: #667eea;
            color: white;
            font-weight: 600;
        }
        
        tr:hover {
            background-color: #f5f5f5;
        }
        
        .table-container {
            overflow-x: auto;
        }
        
        .summary-table th {
            background-color: #764ba2;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 0.9em;
        }
        
        .info-box {
            background-color: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>数据报告</h1>
            <p class="subtitle">生成时间: {{ generate_time }}</p>
        </header>
        
        <div class="section">
            <h2>数据概览</h2>
            <div class="info-box">
                <p>数据行数: {{ data_shape[0] }} | 数据列数: {{ data_shape[1] }}</p>
            </div>
            
            <div class="stats-grid">
                {% for stat in key_stats %}
                <div class="stat-card">
                    <h3>{{ stat.name }}</h3>
                    <div class="value">{{ stat.value }}</div>
                </div>
                {% endfor %}
            </div>
        </div>
        
        <div class="section">
            <h2>图表分析</h2>
            {% for chart in charts %}
            <div class="chart-container">
                <h3>{{ chart.title }}</h3>
                <img src="{{ chart.image }}" alt="{{ chart.title }}">
            </div>
            {% endfor %}
        </div>
        
        <div class="section">
            <h2>数据表格</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            {% for col in data_columns %}
                            <th>{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for _, row in data_head.iterrows() %}
                        <tr>
                            {% for value in row %}
                            <td>{{ value }}</td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                        {% if data_shape[0] > 10 %}
                        <tr>
                            <td colspan="{{ data_shape[1] }}" style="text-align: center; color: #666;">
                                仅显示前10行数据，共{{ data_shape[0] }}行
                            </td>
                        </tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2>统计摘要</h2>
            <div class="table-container">
                <table class="summary-table">
                    <thead>
                        <tr>
                            <th></th>
                            {% for col in summary_columns %}
                            <th>{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for index, row in summary_data.iterrows() %}
                        <tr>
                            <th>{{ index }}</th>
                            {% for value in row %}
                            <td>{{ "{:.4f}".format(value) if value is number else value }}</td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>报告生成器 © {{ year }}</p>
        </div>
    </div>
</body>
</html>"""
        
        # 准备模板数据
        key_stats = []
        if len(numeric_cols) > 0:
            # 添加关键统计指标
            for col in numeric_cols[:4]:  # 只显示前4个数值列的统计
                col_stats = self.data[col].describe()
                key_stats.append({
                    "name": f"{col}平均值",
                    "value": f"{col_stats['mean']:.2f}"
                })
        
        # 如果数值列不足4个，添加其他统计
        if len(key_stats) < 4:
            key_stats.append({
                "name": "数据行数",
                "value": self.data.shape[0]
            })
        
        if len(key_stats) < 4:
            key_stats.append({
                "name": "数据列数",
                "value": self.data.shape[1]
            })
        
        if len(key_stats) < 4:
            key_stats.append({
                "name": "缺失值总数",
                "value": self.data.isnull().sum().sum()
            })
        
        # 渲染HTML
        template = Template(html_template)
        html_content = template.render(
            data_shape=self.data.shape,
            data_columns=self.data.columns.tolist(),
            data_head=self.data.head(10),
            summary_data=self.summary_data,
            summary_columns=self.summary_data.columns.tolist(),
            key_stats=key_stats,
            charts=charts,
            generate_time=time.strftime("%Y-%m-%d %H:%M:%S"),
            year=time.strftime("%Y")
        )
        
        # 保存HTML文件
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        end_time = time.time()
        print(f"HTML生成耗时: {end_time - start_time:.2f}秒")
        
        return output_path
