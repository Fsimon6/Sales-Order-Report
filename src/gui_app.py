import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import logging
import sys
import os
import time

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.excel_source import ExcelSource
from src.excel_mapping import ExcelMapping
from src.feishu_table_service import FeishuTableService


class TextHandler(logging.Handler):
    """自定义日志处理器，将日志输出到Tkinter文本框"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)


class FeishuTableGUI:
    """飞书表格自动化工具 GUI 界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("飞书表格自动化工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        self.running = False
        self.feishu_service = None
        
        # 设置图标（可选）
        try:
            # 尝试设置窗口图标
            pass
        except:
            pass
        
        # 创建组件
        self.create_widgets()
        
        # 配置日志
        self.configure_logging()
        
        # 注册全局停止快捷键
        self.root.bind('<Control-Shift-X>', self.on_stop_hotkey)
    
    def create_widgets(self):
        """创建界面组件"""
        
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(main_frame, text="飞书表格自动化工具", 
                                font=('Microsoft YaHei', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.run_btn = ttk.Button(button_frame, text="开始运行", 
                                   command=self.start_execution,
                                   style='Accent.TButton')
        self.run_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(button_frame, text="停止运行", 
                                   command=self.stop_execution,
                                   state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        
        # 快捷键提示
        hotkey_label = ttk.Label(button_frame, text="停止快捷键: Ctrl+Shift+X", 
                                 font=('Microsoft YaHei', 10))
        hotkey_label.pack(side=tk.RIGHT)
        
        # 操作说明区域
        info_frame = ttk.LabelFrame(main_frame, text="操作说明", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        info_text = """1. 请先打开飞书表格浏览器页面
2. 手动滚动到表格底部，找到最后一行有数据的行
3. 将光标定位到【最后一行数据行的下下行】的【E列】（商品SKU列）
4. 点击"开始运行"按钮或确认后程序将自动执行以下操作：
   - 批量填写SKU到E列
   - 填写商品ID到C列
   - 填写商品名称到D列
   - 填写出单量到H列
   - 填写海外仓库存到I列
5. 运行过程中可使用 Ctrl+Shift+X 快捷键停止操作"""
        
        info_label = ttk.Label(info_frame, text=info_text, justify=tk.LEFT,
                               font=('Microsoft YaHei', 10))
        info_label.pack(fill=tk.X)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, state='disabled',
                                                  font=('Consolas', 10),
                                                  wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                               relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(5, 0))
    
    def configure_logging(self):
        """配置日志输出到文本框和文件"""
        
        # 创建日志处理器
        text_handler = TextHandler(self.log_text)
        text_handler.setLevel(logging.INFO)
        
        # 创建文件处理器
        log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logs')
        os.makedirs(log_dir, exist_ok=True)
        file_handler = logging.FileHandler(os.path.join(log_dir, 'app.log'), 
                                          encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        
        # 设置日志格式
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        text_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)
        
        # 获取根日志记录器
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(text_handler)
        root_logger.addHandler(file_handler)
        
        self.logger = logging.getLogger(__name__)
    
    def start_execution(self):
        """开始执行自动化任务"""
        
        if self.running:
            messagebox.showwarning("提示", "程序正在运行中")
            return
        
        # 清空日志
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        
        self.running = True
        self.run_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("运行中...")
        
        # 在新线程中执行任务
        thread = threading.Thread(target=self.execute_task)
        thread.daemon = True
        thread.start()
    
    def execute_task(self):
        """执行自动化任务的核心逻辑"""
        
        try:
            self.logger.info("========== 开始执行自动化脚本 ==========")
            
            # 1. 读取并处理 Excel 数据
            self.logger.info("Step 1: 读取本地 Excel 文件")
            excel_source = ExcelSource()
            sku_quantity_dict = excel_source.extract_and_aggregate_sku()
            
            if not sku_quantity_dict:
                self.logger.warning("没有提取到任何 SKU 数据")
                self.on_task_complete()
                return
            
            self.logger.info(f"成功提取 {len(sku_quantity_dict)} 个 SKU")
            
            # 2. 打印提取的数据预览
            self.logger.info("提取的数据预览：")
            for sku, qty in list(sku_quantity_dict.items())[:3]:
                self.logger.info(f"  {sku}: {qty}")
            if len(sku_quantity_dict) > 3:
                self.logger.info(f"  ... 还有 {len(sku_quantity_dict) - 3} 个 SKU")
            
            # 3. 查询商品ID映射、商品名称映射和库存映射
            self.logger.info("Step 2: 查询商品ID映射、商品名称映射和库存映射")
            excel_mapping = ExcelMapping()
            
            # 商品ID映射（从总订单文件 → ID映射文件）
            sku_to_product_id_dict = excel_mapping.get_sku_to_product_id_map()
            
            # 商品名称映射（从海外仓库存文件）
            sku_to_product_name_dict = excel_mapping.get_sku_to_product_name_from_warehouse()
            
            # 可用库存映射（从海外仓库存文件）
            sku_to_inventory_dict = excel_mapping.get_sku_to_inventory_map()
            
            if sku_to_product_id_dict:
                self.logger.info(f"成功建立 {len(sku_to_product_id_dict)} 个 SKU → 商品ID 映射")
            else:
                self.logger.warning("未建立任何商品ID映射，将跳过商品ID填写")
            
            if sku_to_product_name_dict:
                self.logger.info(f"成功建立 {len(sku_to_product_name_dict)} 个 SKU → 商品名称 映射")
            else:
                self.logger.warning("未建立任何商品名称映射，将跳过商品名称填写")
            
            if sku_to_inventory_dict:
                self.logger.info(f"成功建立 {len(sku_to_inventory_dict)} 个 SKU → 可用库存 映射")
            else:
                self.logger.warning("未建立任何库存映射，将跳过库存填写")
            
            # 4. 切换到飞书表格并填写数据
            self.logger.info("Step 3: 开始填写飞书表格")
            self.feishu_service = FeishuTableService()
            self.feishu_service.execute(sku_quantity_dict, 
                                        sku_to_product_id_dict, 
                                        sku_to_product_name_dict, 
                                        sku_to_inventory_dict,
                                        skip_confirm=True)
            
            self.logger.info("========== 自动化脚本执行完成 ==========")
            
        except KeyboardInterrupt:
            self.logger.warning("操作被用户中断")
        except Exception as e:
            self.logger.error(f"脚本执行失败: {e}", exc_info=True)
            messagebox.showerror("错误", f"脚本执行失败: {e}")
        finally:
            self.on_task_complete()
    
    def stop_execution(self):
        """停止执行任务"""
        
        if self.feishu_service:
            self.feishu_service.stop()
        self.running = False
        self.logger.warning("正在停止操作...")
    
    def on_stop_hotkey(self, event):
        """响应停止快捷键"""
        
        if self.running:
            self.stop_execution()
    
    def on_task_complete(self):
        """任务完成后的清理工作"""
        
        self.running = False
        self.run_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")
        
        if self.feishu_service:
            # 移除快捷键监听
            try:
                import keyboard
                keyboard.remove_hotkey('ctrl+shift+x')
            except:
                pass
            self.feishu_service = None


def main():
    """GUI程序入口"""
    
    root = tk.Tk()
    app = FeishuTableGUI(root)
    
    # 窗口关闭时的处理
    def on_closing():
        if app.running:
            if messagebox.askokcancel("退出", "程序正在运行中，确定要退出吗？"):
                app.stop_execution()
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # 启动主循环
    root.mainloop()


if __name__ == '__main__':
    main()
