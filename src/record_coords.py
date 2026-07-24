"""
坐标记录脚本

用于记录飞书表格Ctrl+F搜索框中左右箭头按钮的坐标

使用方法：
1. 打开飞书表格
2. 按Ctrl+F打开查找对话框
3. 运行此脚本
4. 将鼠标移动到左箭头按钮上，按空格键记录坐标
5. 将鼠标移动到右箭头按钮上，按空格键记录坐标
6. 按ESC退出脚本
7. 坐标会保存到 coords.json 文件中
"""

import pyautogui
import keyboard
import json
import os

def main():
    print("="*60)
    print("坐标记录脚本")
    print("="*60)
    print("使用说明：")
    print("1. 打开飞书表格")
    print("2. 按Ctrl+F打开查找对话框")
    print("3. 将鼠标移动到左箭头按钮上，按空格键记录坐标")
    print("4. 将鼠标移动到右箭头按钮上，按空格键记录坐标")
    print("5. 按ESC退出脚本")
    print("="*60)
    
    coords = {}
    
    # 记录左箭头坐标
    print("\n请将鼠标移动到【左箭头】按钮上，按空格键记录坐标...")
    keyboard.wait('space')
    left_arrow_x, left_arrow_y = pyautogui.position()
    coords['left_arrow'] = {'x': left_arrow_x, 'y': left_arrow_y}
    print(f"左箭头坐标: ({left_arrow_x}, {left_arrow_y})")
    
    # 记录右箭头坐标
    print("\n请将鼠标移动到【右箭头】按钮上，按空格键记录坐标...")
    keyboard.wait('space')
    right_arrow_x, right_arrow_y = pyautogui.position()
    coords['right_arrow'] = {'x': right_arrow_x, 'y': right_arrow_y}
    print(f"右箭头坐标: ({right_arrow_x}, {right_arrow_y})")
    
    # 保存坐标
    save_path = os.path.join(os.path.dirname(__file__), 'coords.json')
    with open(save_path, 'w', encoding='utf-8') as f:
        json.dump(coords, f, indent=2, ensure_ascii=False)
    
    print(f"\n坐标已保存到: {save_path}")
    print("按ESC退出...")
    keyboard.wait('esc')

if __name__ == '__main__':
    main()


