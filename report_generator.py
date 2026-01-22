#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel报表生成器
输入数据→3秒生成带图表报告 + index.html
"""

import pandas as pd
import time
import argparse
from data_processor import DataProcessor
from excel_generator import ExcelGenerator
from html_generator import HTMLGenerator

class ReportGenerator:
    def __init__(self):
        self.data_processor = DataProcessor()
        self.excel_generator = None
        self.html_generator = HTMLGenerator()
    
    def generate_report(self, input_data, excel_output="report.xlsx", html_output="index.html"):
        """
        生成完整的报告
        :param input_data: 输入数据，可以是文件路径、字典、列表或DataFrame
        :param excel_output: Excel输出路径
        :param html_output: HTML输出路径
        :return: 生成的文件路径列表
        """
        start_time = time.time()
        
        print("正在处理数据...")
        # 1. 处理数据
        data = self.data_processor.load_data(input_data)
        
        print("正在生成Excel报告...")
        # 2. 生成Excel报告
        self.excel_generator = ExcelGenerator()
        excel_path = self.excel_generator.generate_excel(data, excel_output)
        
        print("正在生成HTML报告...")
        # 3. 生成HTML报告
        self.html_generator.set_data(data)
        html_path = self.html_generator.generate_html(html_output)
        
        total_time = time.time() - start_time
        print(f"\n报告生成完成！总耗时: {total_time:.2f}秒")
        print(f"Excel报告: {excel_path}")
        print(f"HTML报告: {html_path}")
        
        if total_time <= 3:
            print("✅ 符合3秒内完成的要求！")
        else:
            print("⚠️  生成时间超过3秒，建议优化数据量或配置")
        
        return [excel_path, html_path]

def main():
    parser = argparse.ArgumentParser(description="Excel报表生成器")
    parser.add_argument("-i", "--input", type=str, help="输入数据文件路径 (CSV/Excel/JSON)")
    parser.add_argument("-e", "--excel", type=str, default="report.xlsx", help="Excel输出路径")
    parser.add_argument("-t", "--html", type=str, default="index.html", help="HTML输出路径")
    parser.add_argument("--demo", action="store_true", help="使用演示数据生成报告")
    
    args = parser.parse_args()
    
    generator = ReportGenerator()
    
    if args.demo:
        # 使用演示数据
        print("正在生成演示数据...")
        demo_data = {
            "月份": ["1月", "2月", "3月", "4月", "5月", "6月"],
            "销售额": [12000, 15000, 18000, 16000, 20000, 22000],
            "利润": [3600, 4500, 5400, 4800, 6000, 6600],
            "客户数量": [120, 150, 180, 160, 200, 220]
        }
        generator.generate_report(demo_data, args.excel, args.html)
    elif args.input:
        # 从文件加载数据
        generator.generate_report(args.input, args.excel, args.html)
    else:
        # 交互式模式
        print("欢迎使用Excel报表生成器！")
        print("请选择输入方式:")
        print("1. 从文件加载数据 (CSV/Excel/JSON)")
        print("2. 使用演示数据")
        
        choice = input("请输入选项 (1/2): ").strip()
        
        if choice == "1":
            file_path = input("请输入文件路径: ").strip()
            generator.generate_report(file_path, args.excel, args.html)
        elif choice == "2":
            demo_data = {
                "月份": ["1月", "2月", "3月", "4月", "5月", "6月"],
                "销售额": [12000, 15000, 18000, 16000, 20000, 22000],
                "利润": [3600, 4500, 5400, 4800, 6000, 6600],
                "客户数量": [120, 150, 180, 160, 200, 220]
            }
            generator.generate_report(demo_data, args.excel, args.html)
        else:
            print("无效选项，程序退出")

if __name__ == "__main__":
    main()
