import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()


class ExcelSource:
    """
    读取本地 Excel 文件，提取 Seller SKU 和 Quantity 数据，去重并汇总
    """

    def __init__(self):
        self.source_path = os.getenv('EXCEL_SOURCE_PATH', 'excel_source/')

    def get_latest_excel_file(self):
        """获取最新的 Excel 文件（优先查找包含'总'字的订单文件）"""
        files = [f for f in os.listdir(self.source_path)
                 if f.endswith('.xlsx') and not f.startswith('~$')]
        if not files:
            raise FileNotFoundError(f"在 {self.source_path} 目录下未找到 Excel 文件")

        # 优先查找包含'总'字的订单文件
        order_files = [f for f in files if '总' in f]
        if order_files:
            files_with_time = [(f, os.path.getmtime(os.path.join(self.source_path, f)))
                               for f in order_files]
            files_with_time.sort(key=lambda x: x[1], reverse=True)
            return os.path.join(self.source_path, files_with_time[0][0])

        # 如果没有找到'总'字文件，使用最新修改的文件
        files_with_time = [(f, os.path.getmtime(os.path.join(self.source_path, f)))
                           for f in files]
        files_with_time.sort(key=lambda x: x[1], reverse=True)
        return os.path.join(self.source_path, files_with_time[0][0])

    def extract_and_aggregate_sku(self, file_path=None):
        """
        提取 Seller SKU 和 Quantity 列，去重并汇总
        """
        if file_path is None:
            file_path = self.get_latest_excel_file()

        print(f"正在读取文件: {file_path}")

        # 读取 Excel 文件，第一行是表头，跳过第二行（标题说明行）
        df = pd.read_excel(file_path, skiprows=[1])

        # 查找 SKU 列（可能的列名）
        sku_column = None
        for col in df.columns:
            if 'sku' in str(col).lower() and 'seller' in str(col).lower():
                sku_column = col
                break

        if sku_column is None:
            for col in df.columns:
                if 'sku' in str(col).lower():
                    sku_column = col
                    break

        if sku_column is None:
            raise ValueError("Excel 文件中未找到 SKU 列")

        # 查找 Quantity 列（可能的列名）
        quantity_column = None
        for col in df.columns:
            if 'quantity' in str(col).lower() and 'sold' in str(col).lower():
                quantity_column = col
                break

        if quantity_column is None:
            for col in df.columns:
                if 'quantity' in str(col).lower():
                    quantity_column = col
                    break

        if quantity_column is None:
            raise ValueError("Excel 文件中未找到 Quantity 列")

        print(f"找到 SKU 列: {sku_column}")
        print(f"找到 Quantity 列: {quantity_column}")

        # 提取相关列
        sku_data = df[[sku_column, quantity_column]].copy()

        # 去除空值
        sku_data = sku_data.dropna(subset=[sku_column])

        # 将 Quantity 转换为数值类型，非数值转为0
        sku_data[quantity_column] = pd.to_numeric(sku_data[quantity_column], errors='coerce').fillna(0)

        # 按 SKU 汇总 Quantity
        aggregated = sku_data.groupby(sku_column)[quantity_column].sum().astype(int).to_dict()

        print(f"共提取 {len(aggregated)} 个唯一 SKU")
        return aggregated


if __name__ == '__main__':
    excel_source = ExcelSource()
    sku_quantity_dict = excel_source.extract_and_aggregate_sku()
    print("SKU 汇总结果:")
    for sku, qty in list(sku_quantity_dict.items())[:5]:
        print(f"  {sku}: {qty}")
    print(f"  ... 共 {len(sku_quantity_dict)} 条")
