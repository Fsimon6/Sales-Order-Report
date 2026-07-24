import logging
import argparse
from src.excel_source import ExcelSource
from src.excel_mapping import ExcelMapping
from src.feishu_table_service import FeishuTableService

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def main():
    """
    主程序入口
    流程：
    1. 从本地 Excel 读取并汇总 SKU 数据
    2. 查询商品ID映射
    3. 切换到飞书表格，填写数据
    """
    parser = argparse.ArgumentParser(description='飞书表格自动化脚本')
    parser.add_argument('-s', '--skip-confirm', action='store_true', help='跳过确认')
    args = parser.parse_args()

    logger.info("========== 开始执行自动化脚本 ==========")

    try:
        # 1. 读取并处理 Excel 数据
        logger.info("Step 1: 读取本地 Excel 文件")
        excel_source = ExcelSource()
        sku_quantity_dict = excel_source.extract_and_aggregate_sku()

        if not sku_quantity_dict:
            logger.warning("没有提取到任何 SKU 数据")
            return

        logger.info(f"成功提取 {len(sku_quantity_dict)} 个 SKU")

        # 2. 打印提取的数据预览
        logger.info("提取的数据预览：")
        for sku, qty in list(sku_quantity_dict.items())[:3]:
            logger.info(f"  {sku}: {qty}")
        if len(sku_quantity_dict) > 3:
            logger.info(f"  ... 还有 {len(sku_quantity_dict) - 3} 个 SKU")

        # 3. 查询商品ID映射、商品名称映射（从海外仓库存文件）和库存映射
        logger.info("Step 2: 查询商品ID映射、商品名称映射和库存映射")
        excel_mapping = ExcelMapping()
        
        # 商品ID映射（从总订单文件 → ID映射文件）
        sku_to_product_id_dict = excel_mapping.get_sku_to_product_id_map()
        
        # 商品名称映射（从海外仓库存文件）
        sku_to_product_name_dict = excel_mapping.get_sku_to_product_name_from_warehouse()
        
        # 可用库存映射（从海外仓库存文件，支持多行求和）
        sku_to_inventory_dict = excel_mapping.get_sku_to_inventory_map()
        
        if sku_to_product_id_dict:
            logger.info(f"成功建立 {len(sku_to_product_id_dict)} 个 SKU → 商品ID 映射")
        else:
            logger.warning("未建立任何商品ID映射，将跳过商品ID填写")
            
        if sku_to_product_name_dict:
            logger.info(f"成功建立 {len(sku_to_product_name_dict)} 个 SKU → 商品名称 映射")
        else:
            logger.warning("未建立任何商品名称映射，将跳过商品名称填写")
            
        if sku_to_inventory_dict:
            logger.info(f"成功建立 {len(sku_to_inventory_dict)} 个 SKU → 可用库存 映射")
        else:
            logger.warning("未建立任何库存映射，将跳过库存填写")

        # 4. 切换到飞书表格并填写数据
        logger.info("Step 3: 开始填写飞书表格")
        feishu_service = FeishuTableService()
        feishu_service.execute(sku_quantity_dict, sku_to_product_id_dict, sku_to_product_name_dict, sku_to_inventory_dict, skip_confirm=args.skip_confirm)

        logger.info("========== 自动化脚本执行完成 ==========")
        print("\n操作完成！")

    except Exception as e:
        logger.error(f"脚本执行失败: {e}", exc_info=True)
        print(f"\n操作失败: {e}")


if __name__ == '__main__':
    main()
