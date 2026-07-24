# 飞书表格自动化工具

自动读取Excel订单数据，填写到飞书表格中。

## 功能特性

- 自动读取Excel文件中的订单数据
- 批量填写商品SKU、商品ID、商品名称、出单量、海外仓库存到飞书表格
- 支持大小写不敏感的SKU查询
- 提供GUI可视化界面和运行日志
- 支持快捷键停止操作（Ctrl+Shift+X）

## 技术栈

- Python 3.10+
- PyAutoGUI - GUI自动化
- pandas - Excel数据处理
- Tkinter - GUI界面
- pywin32 - 剪贴板操作

## 项目结构

```
.
├── main.py                 # 主程序入口
├── requirements.txt        # 依赖列表
├── README.md               # 项目说明
├── .gitignore              # 忽略规则
└── src/
    ├── excel_source.py     # Excel数据读取模块
    ├── excel_mapping.py    # SKU映射逻辑模块
    ├── feishu_table_service.py  # 飞书表格操作模块
    ├── gui_app.py          # GUI界面模块
    └── record_coords.py    # 坐标录制工具
```

## 使用方法

### 开发环境运行

```bash
# 创建虚拟环境
python -m venv .venv

# 激活虚拟环境
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 运行GUI
python src/gui_app.py
```

### 打包成可执行文件

```bash
pyinstaller feishu_table.spec
```

生成的exe文件位于 `dist/` 目录。

## Excel文件要求

将以下文件放入 `excel_source/` 目录：

| 文件类型 | 文件名要求 | 说明 |
|---------|-----------|------|
| 订单文件 | 包含"总"字 | 如"7.21总订单.xlsx" |
| ID映射文件 | 包含"ID"字 | 如"跨境2店ID.xlsx" |
| 库存文件 | 包含"海外仓"字 | 如"海外仓库存.xlsx" |

## 操作步骤

1. 打开飞书表格网页，将光标定位到目标单元格
2. 运行程序，点击"开始运行"按钮
3. 等待程序自动填写数据

## 停止操作

- 快捷键：Ctrl + Shift + X
- 或点击界面上的"停止运行"按钮
