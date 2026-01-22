import pandas as pd
import json
from typing import Union, Dict, List, Any

class DataProcessor:
    def __init__(self):
        self.data = None
        
    def load_data(self, input_data: Union[str, Dict, List, pd.DataFrame]) -> pd.DataFrame:
        """
        支持多种数据输入格式
        :param input_data: 可以是文件路径(str)、字典(Dict)、列表(List)或DataFrame
        :return: 处理后的DataFrame
        """
        if isinstance(input_data, pd.DataFrame):
            self.data = input_data
        elif isinstance(input_data, str):
            self._load_from_file(input_data)
        elif isinstance(input_data, (dict, list)):
            self._load_from_structure(input_data)
        else:
            raise ValueError("不支持的数据格式，请提供文件路径、字典、列表或DataFrame")
        
        return self.data
    
    def _load_from_file(self, file_path: str) -> None:
        """从文件加载数据"""
        if file_path.endswith('.csv'):
            self.data = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
            self.data = pd.read_excel(file_path)
        elif file_path.endswith('.json'):
            with open(file_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            self._load_from_structure(json_data)
        else:
            raise ValueError("不支持的文件格式，请提供CSV、Excel或JSON文件")
    
    def _load_from_structure(self, data: Union[Dict, List]) -> None:
        """从数据结构加载数据"""
        if isinstance(data, dict):
            # 支持两种字典格式：{"columns": [...], "data": [...]} 或直接的{"列名": [值列表]}
            if "columns" in data and "data" in data:
                self.data = pd.DataFrame(data["data"], columns=data["columns"])
            else:
                self.data = pd.DataFrame(data)
        elif isinstance(data, list):
            # 列表格式：[{"列名": 值}, {"列名": 值}, ...] 或 [[值1, 值2, ...], [值1, 值2, ...]]
            if all(isinstance(item, dict) for item in data):
                self.data = pd.DataFrame(data)
            else:
                self.data = pd.DataFrame(data)
    
    def get_summary_statistics(self) -> pd.DataFrame:
        """获取数据的基本统计信息"""
        if self.data is None:
            raise ValueError("没有加载数据，请先调用load_data方法")
        return self.data.describe()
    
    def get_data_types(self) -> pd.Series:
        """获取各列的数据类型"""
        if self.data is None:
            raise ValueError("没有加载数据，请先调用load_data方法")
        return self.data.dtypes
    
    def get_column_names(self) -> List[str]:
        """获取列名列表"""
        if self.data is None:
            raise ValueError("没有加载数据，请先调用load_data方法")
        return list(self.data.columns)
