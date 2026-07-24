import time
import pyautogui
import keyboard
import logging
import win32clipboard
import win32con
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class Clipboard:
    """剪贴板操作工具"""
    
    @staticmethod
    def set_text(text):
        """设置剪贴板文本"""
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    
    @staticmethod
    def get_text():
        """获取剪贴板文本"""
        win32clipboard.OpenClipboard()
        try:
            text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        except:
            text = ""
        win32clipboard.CloseClipboard()
        return text


class FeishuTableService:
    """
    飞书表格自动化操作服务
    
    列对应关系：
    C列 - 商品ID
    D列 - 商品名称
    E列 - 商品SKU
    H列 - 出单量
    I列 - 海外仓剩余货量（美东＋美西）
    """

    def __init__(self):
        self.running = True
        keyboard.add_hotkey('ctrl+shift+x', self.stop)

    def stop(self):
        """停止执行（全局快捷键 Ctrl+Shift+X）"""
        logger.warning("检测到停止信号，正在停止操作...")
        self.running = False

    def wait(self, seconds=0.3):
        """等待指定时间，同时检查是否需要停止"""
        for _ in range(int(seconds * 10)):
            if not self.running:
                raise KeyboardInterrupt("用户通过快捷键停止操作")
            time.sleep(0.1)

    def switch_to_chrome(self):
        """切换到 Chrome 浏览器窗口"""
        logger.info("切换到 Chrome 浏览器...")
        pyautogui.keyDown('alt')
        pyautogui.press('tab')
        pyautogui.keyUp('alt')
        self.wait(2)
        
        # 点击表格区域激活焦点
        logger.info("激活表格焦点...")
        pyautogui.click()
        self.wait(1)

    def paste_text(self, text):
        """将文本粘贴到当前单元格"""
        Clipboard.set_text(text)
        self.wait(0.2)
        pyautogui.hotkey('ctrl', 'v')
        self.wait(0.3)

    def move_right(self, steps=1):
        """向右移动"""
        for _ in range(steps):
            pyautogui.press('right')
            self.wait(0.1)

    def move_left(self, steps=1):
        """向左移动"""
        for _ in range(steps):
            pyautogui.press('left')
            self.wait(0.1)

    def move_down(self, steps=1):
        """向下移动"""
        for _ in range(steps):
            pyautogui.press('down')
            self.wait(0.1)

    def move_up(self, steps=1):
        """向上移动"""
        for _ in range(steps):
            pyautogui.press('up')
            self.wait(0.1)

    def execute(self, sku_quantity_dict, sku_to_product_id_dict=None, sku_to_product_name_dict=None, sku_to_inventory_dict=None, skip_confirm=False):
        """
        执行完整的飞书表格操作流程
        
        用户需要在运行前手动定位到目标空行的E列（商品SKU列）
        
        流程：
        1. 批量填写所有SKU到E列（连续填写，复制粘贴）
        2. 回到第一个SKU行
        3. 逐个填写商品ID到C列和商品名称到D列
        4. 逐个填写出单量到H列
        5. 逐个填写海外仓库存到I列
        
        :param sku_quantity_dict: SKU到出单量的映射
        :param sku_to_product_id_dict: SKU到商品ID的映射（可选）
        :param sku_to_product_name_dict: SKU到商品名称的映射（可选）
        :param sku_to_inventory_dict: SKU到可用库存的映射（可选）
        :param skip_confirm: 是否跳过确认
        """
        if not sku_quantity_dict:
            logger.warning("没有数据需要处理")
            return

        try:
            # 显示操作说明
            print("\n" + "="*60)
            print("操作说明：")
            print("1. 请先打开飞书表格")
            print("2. 手动滚动到表格底部")
            print("3. 找到最后一行有数据的行")
            print("4. 将光标定位到【最后一行数据行的下下行】的【E列】（商品SKU列）")
            print("5. 确认光标已正确定位后，按回车键继续")
            print("="*60)
            if not skip_confirm:
                input("按回车键继续...")

            # 切换到浏览器并激活表格
            self.switch_to_chrome()
            self.wait(2)

            sku_list = list(sku_quantity_dict.keys())
            total_sku = len(sku_list)

            # ====== 步骤1：批量填写所有SKU到E列 ======
            logger.info(f"步骤1: 批量填写 {total_sku} 个 SKU...")
            
            for idx, sku in enumerate(sku_list):
                if not self.running:
                    break
                
                logger.info(f"粘贴 SKU {idx + 1}/{total_sku}: {sku}")
                self.paste_text(sku)
                
                # 向下移动1行（连续填写）
                if idx < total_sku - 1:
                    self.move_down(1)
            
            logger.info("SKU 填写完成")

            # ====== 步骤2：回到第一个SKU行 ======
            rows_to_move_up = total_sku - 1
            logger.info(f"步骤2: 向上移动 {rows_to_move_up} 行回到第一个SKU")
            self.move_up(rows_to_move_up)
            self.wait(0.5)

            # ====== 步骤3：逐个填写商品ID到C列和商品名称到D列 ======
            if sku_to_product_id_dict:
                logger.info("步骤3: 逐个填写商品ID和商品名称...")
                for idx, sku in enumerate(sku_list):
                    if not self.running:
                        break

                    logger.info(f"处理 SKU {idx + 1}/{total_sku}: {sku}")
                    
                    # 当前在E列，向左移动2步到C列（E→D→C）
                    self.move_left(2)
                    
                    # 填写商品ID
                    product_id = sku_to_product_id_dict.get(sku, "")
                    if product_id:
                        logger.info(f"填写商品ID到C列: {product_id}")
                        self.paste_text(product_id)
                    else:
                        logger.warning(f"SKU {sku} 未找到对应的商品ID，跳过")
                    
                    # 移动到D列（向右1步）
                    self.move_right(1)
                    
                    # 填写商品名称（支持大小写不敏感查询）
                    product_name = sku_to_product_name_dict.get(sku, "")
                    if not product_name:
                        # 尝试小写SKU查询
                        sku_lower = sku.lower()
                        product_name = sku_to_product_name_dict.get(sku_lower, "")
                        if product_name:
                            logger.info(f"通过小写SKU ({sku_lower}) 找到商品名称: {product_name[:30]}...")
                    if product_name:
                        logger.info(f"填写商品名称到D列: {product_name[:30]}...")
                        self.paste_text(product_name)
                    else:
                        logger.warning(f"SKU {sku} 未找到对应的商品名称，跳过")
                    
                    # 回到E列，准备处理下一个SKU
                    # 当前在D列，向右移动1步到E列（D→E）
                    self.move_right(1)
                    
                    # 移动到下一个SKU行（向下1行）
                    if idx < total_sku - 1:
                        self.move_down(1)

                # ====== 回到第一个SKU行 ======
                # 由于最后一行没有执行move_down，所以只需要向上移动 total_sku - 1 行
                rows_to_move_up = total_sku - 1
                logger.info(f"步骤3完成: 向上移动 {rows_to_move_up} 行回到第一个SKU")
                self.move_up(rows_to_move_up)
                self.wait(0.5)

            # ====== 步骤4：逐个填写出单量到H列 ======
            logger.info("步骤4: 逐个填写出单量...")
            for idx, sku in enumerate(sku_list):
                if not self.running:
                    break

                logger.info(f"处理 SKU {idx + 1}/{total_sku}: {sku}")
                
                # 当前在E列，向右移动3步到H列（E→F→G→H）
                self.move_right(3)
                
                # 填写出单量
                quantity = sku_quantity_dict[sku]
                logger.info(f"填写出单量到H列: {quantity}")
                self.paste_text(str(quantity))
                
                # 回到E列，准备处理下一个SKU
                # 当前在H列，向左移动3步到E列（H→G→F→E）
                self.move_left(3)
                
                # 移动到下一个SKU行（向下1行）
                if idx < total_sku - 1:
                    self.move_down(1)

            # ====== 回到第一个SKU行 ======
            rows_to_move_up = total_sku - 1
            logger.info(f"步骤4完成: 向上移动 {rows_to_move_up} 行回到第一个SKU")
            self.move_up(rows_to_move_up)
            self.wait(0.5)

            # ====== 步骤5：逐个填写海外仓库存到I列 ======
            if sku_to_inventory_dict:
                logger.info("步骤5: 逐个填写海外仓库存...")
                for idx, sku in enumerate(sku_list):
                    if not self.running:
                        break

                    logger.info(f"处理 SKU {idx + 1}/{total_sku}: {sku}")
                    
                    # 当前在E列，向右移动4步到I列（E→F→G→H→I）
                    self.move_right(4)
                    
                    # 填写库存（支持大小写不敏感查询，库存为0也需填写）
                    inventory = sku_to_inventory_dict.get(sku, "")
                    if inventory is None or inventory == "":
                        # 尝试小写SKU查询
                        sku_lower = sku.lower()
                        inventory = sku_to_inventory_dict.get(sku_lower, "")
                        if inventory is not None and inventory != "":
                            logger.info(f"通过小写SKU ({sku_lower}) 找到库存: {inventory}")
                    if inventory is not None and inventory != "":
                        logger.info(f"填写海外仓库存到I列: {inventory}")
                        self.paste_text(str(inventory))
                    else:
                        logger.warning(f"SKU {sku} 未找到对应的库存，跳过")
                    
                    # 回到E列，准备处理下一个SKU
                    # 当前在I列，向左移动4步到E列（I→H→G→F→E）
                    self.move_left(4)
                    
                    # 移动到下一个SKU行（向下1行）
                    if idx < total_sku - 1:
                        self.move_down(1)

            logger.info("所有操作完成")

        except KeyboardInterrupt:
            logger.warning("操作被用户中断")
        except Exception as e:
            logger.error(f"操作失败: {e}", exc_info=True)
            raise


if __name__ == '__main__':
    service = FeishuTableService()
    test_data = {
        'SKU001': 10,
        'SKU002': 5,
        'SKU003': 8
    }
    service.execute(test_data)
