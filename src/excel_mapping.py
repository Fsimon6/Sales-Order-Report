import os
import pandas as pd
import logging

logger = logging.getLogger(__name__)


class ExcelMapping:
    """
    Excel映射服务
    
    用于处理商品ID的查询：
    1. 从总订单文件中通过SKU查找Product Name
    2. 从ID映射文件中通过Product Name查找商品ID
    3. 从海外仓库存文件中通过SKU查找Product Name和可用库存（支持多行求和）
    """

    def __init__(self, source_path='excel_source/'):
        self.source_path = source_path

    def find_file_by_keyword(self, keyword):
        """
        通过关键字查找Excel文件
        
        :param keyword: 关键字（如"总"、"ID"）
        :return: 文件路径，如果未找到返回None
        """
        files = [f for f in os.listdir(self.source_path)
                 if f.endswith('.xlsx') and not f.startswith('~$')]
        
        for f in files:
            if keyword in f:
                return os.path.join(self.source_path, f)
        
        return None

    def get_sku_to_product_name_map(self):
        """
        从总订单文件中建立 SKU → Product Name 的映射
        
        :return: 字典 {sku: product_name}
        """
        file_path = self.find_file_by_keyword('总')
        if not file_path:
            logger.error("未找到包含'总'字的总订单文件")
            return {}

        logger.info(f"正在读取总订单文件: {file_path}")

        # 读取Excel，第一行是表头，跳过第二行（标题说明行）
        df = pd.read_excel(file_path, skiprows=[1])

        # 查找Seller SKU列
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
            logger.error("总订单文件中未找到SKU列")
            return {}

        # 查找Product Name列
        product_name_column = None
        for col in df.columns:
            if 'product' in str(col).lower() and 'name' in str(col).lower():
                product_name_column = col
                break

        if product_name_column is None:
            logger.error("总订单文件中未找到Product Name列")
            return {}

        logger.info(f"找到 SKU 列: {sku_column}")
        logger.info(f"找到 Product Name 列: {product_name_column}")

        # 提取数据，去重（保留第一个匹配）
        sku_to_product = {}
        for _, row in df.iterrows():
            sku = str(row[sku_column]).strip()
            product_name = str(row[product_name_column]).strip()
            
            if sku and product_name and sku not in sku_to_product:
                sku_to_product[sku] = product_name

        logger.info(f"共建立 {len(sku_to_product)} 个 SKU → Product Name 映射")
        return sku_to_product

    def get_product_name_to_id_map(self):
        """
        从ID映射文件中建立 Product Name → 商品ID 的映射
        
        :return: 字典 {product_name: product_id}
        """
        file_path = self.find_file_by_keyword('ID')
        if not file_path:
            logger.error("未找到包含'ID'字的ID映射文件")
            return {}

        logger.info(f"正在读取ID映射文件: {file_path}")

        # ID映射文件的第4行（索引3）是真正的表头
        df = pd.read_excel(file_path, header=3)

        # 查找商品名列
        product_name_column = None
        for col in df.columns:
            if '商品名' in str(col):
                product_name_column = col
                break

        if product_name_column is None:
            logger.error("ID映射文件中未找到商品名列")
            return {}

        # 查找商品ID列
        product_id_column = None
        for col in df.columns:
            if '商品 ID' in str(col) or '商品ID' in str(col):
                product_id_column = col
                break

        if product_id_column is None:
            logger.error("ID映射文件中未找到商品ID列")
            return {}

        logger.info(f"找到 商品名 列: {product_name_column}")
        logger.info(f"找到 商品ID 列: {product_id_column}")

        # 提取数据，去重（保留第一个匹配）
        product_to_id = {}
        for _, row in df.iterrows():
            product_name = str(row[product_name_column]).strip()
            product_id = str(row[product_id_column]).strip()
            
            if product_name and product_id and product_name not in product_to_id:
                product_to_id[product_name] = product_id

        logger.info(f"共建立 {len(product_to_id)} 个 Product Name → 商品ID 映射")
        return product_to_id

    def get_sku_to_product_name_map_only(self):
        """
        只获取 SKU → Product Name 的映射（用于填写D列商品名称）
        
        :return: 字典 {sku: product_name}
        """
        return self.get_sku_to_product_name_map()

    def get_sku_to_product_name_from_warehouse(self):
        """
        从海外仓库存文件中建立 SKU → Product Name 的映射
        
        支持大小写不敏感查询：当使用原始SKU查询不到时，会尝试使用小写SKU查询
        
        :return: 字典 {sku: product_name}
        """
        file_path = self.find_file_by_keyword('海外仓')
        if not file_path:
            logger.error("未找到包含'海外仓'字的库存文件")
            return {}

        logger.info(f"正在读取海外仓库存文件: {file_path}")

        # 读取Excel
        df = pd.read_excel(file_path)

        # 查找SKU列
        sku_column = None
        for col in df.columns:
            if 'sku' in str(col).lower():
                sku_column = col
                break

        if sku_column is None:
            logger.error("海外仓库存文件中未找到SKU列")
            return {}

        # 查找Product Name列
        product_name_column = None
        for col in df.columns:
            if 'product name' in str(col).lower() or '产品名称' in str(col):
                product_name_column = col
                break

        if product_name_column is None:
            logger.error("海外仓库存文件中未找到Product Name/产品名称列")
            return {}

        logger.info(f"找到 SKU 列: {sku_column}")
        logger.info(f"找到 Product Name 列: {product_name_column}")

        # 提取数据，去重（保留第一个匹配）
        # 同时建立原始SKU和小写SKU的映射，支持大小写不敏感查询
        sku_to_product = {}
        sku_to_product_lower = {}  # 小写SKU映射
        
        for _, row in df.iterrows():
            sku = str(row[sku_column]).strip()
            product_name = str(row[product_name_column]).strip()
            
            if sku and product_name:
                if sku not in sku_to_product:
                    sku_to_product[sku] = product_name
                # 建立小写映射（用于回退查询）
                sku_lower = sku.lower()
                if sku_lower not in sku_to_product_lower:
                    sku_to_product_lower[sku_lower] = product_name

        # 将小写映射合并到主映射中（仅添加原始映射中没有的）
        for sku_lower, product_name in sku_to_product_lower.items():
            if sku_lower not in sku_to_product:
                sku_to_product[sku_lower] = product_name

        logger.info(f"共建立 {len(sku_to_product)} 个 SKU → Product Name 映射（含小写SKU回退）")
        return sku_to_product

    def get_sku_to_inventory_map(self):
        """
        从海外仓库存文件中建立 SKU → 可用库存总和 的映射
        如果一个SKU对应多行数据，将可用库存求和
        
        支持大小写不敏感查询：当使用原始SKU查询不到时，会尝试使用小写SKU查询
        
        :return: 字典 {sku: total_inventory}
        """
        file_path = self.find_file_by_keyword('海外仓')
        if not file_path:
            logger.error("未找到包含'海外仓'字的库存文件")
            return {}

        logger.info(f"正在读取海外仓库存文件: {file_path}")

        # 读取Excel
        df = pd.read_excel(file_path)

        # 查找SKU列
        sku_column = None
        for col in df.columns:
            if 'sku' in str(col).lower():
                sku_column = col
                break

        if sku_column is None:
            logger.error("海外仓库存文件中未找到SKU列")
            return {}

        # 查找可用库存列
        inventory_column = None
        for col in df.columns:
            if 'available' in str(col).lower() or '可用库存' in str(col):
                inventory_column = col
                break

        if inventory_column is None:
            logger.error("海外仓库存文件中未找到Available Inventory/可用库存列")
            return {}

        logger.info(f"找到 SKU 列: {sku_column}")
        logger.info(f"找到 可用库存 列: {inventory_column}")

        # 先按SKU分组求和（使用逐行处理，与商品名称查询保持一致）
        sku_to_inventory = {}
        
        # 第一步：逐行读取，按SKU累加库存
        for _, row in df.iterrows():
            sku = str(row[sku_column]).strip()
            inventory_value = row[inventory_column]
            
            if sku:
                # 将库存值转换为整数
                try:
                    inventory_int = int(float(inventory_value)) if pd.notna(inventory_value) else 0
                except:
                    inventory_int = 0
                
                # 累加库存
                if sku in sku_to_inventory:
                    sku_to_inventory[sku] += inventory_int
                else:
                    sku_to_inventory[sku] = inventory_int

        # 第二步：建立小写SKU回退映射
        for sku in list(sku_to_inventory.keys()):
            sku_lower = sku.lower()
            if sku_lower != sku and sku_lower not in sku_to_inventory:
                sku_to_inventory[sku_lower] = sku_to_inventory[sku]

        logger.info(f"共建立 {len(sku_to_inventory)} 个 SKU → 可用库存 映射（含小写SKU回退）")
        return sku_to_inventory

    def get_sku_to_product_id_map(self):
        """
        组合两个映射，建立 SKU → 商品ID 的映射
        
        :return: 字典 {sku: product_id}
        """
        logger.info("开始建立SKU到商品ID的映射...")

        # 获取两个映射
        sku_to_product = self.get_sku_to_product_name_map()
        product_to_id = self.get_product_name_to_id_map()

        # 组合映射
        sku_to_id = {}
        not_found_sku = []

        for sku, product_name in sku_to_product.items():
            if product_name in product_to_id:
                sku_to_id[sku] = product_to_id[product_name]
                logger.debug(f"SKU {sku} -> Product Name '{product_name}' -> 商品ID {product_to_id[product_name]}")
            else:
                not_found_sku.append((sku, product_name))
                logger.warning(f"SKU {sku} 的 Product Name '{product_name}' 在ID映射文件中未找到")

        logger.info(f"共建立 {len(sku_to_id)} 个 SKU → 商品ID 映射")
        if not_found_sku:
            logger.info(f"有 {len(not_found_sku)} 个SKU未找到对应的商品ID")

        return sku_to_id


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    mapping = ExcelMapping()
    sku_to_id = mapping.get_sku_to_product_id_map()
    
    print("\nSKU → 商品ID 映射结果:")
    for sku, product_id in list(sku_to_id.items())[:10]:
        print(f"  {sku}: {product_id}")
    print(f"  ... 共 {len(sku_to_id)} 条")
