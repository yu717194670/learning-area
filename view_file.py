#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分页查看文件内容，支持翻页并记录当前位置
"""

import os
import sys

def view_file(file_path, page_size=20):
    """
    分页查看文件内容
    :param file_path: 文件路径
    :param page_size: 每页显示的行数
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return
    
    # 读取文件内容
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    total_lines = len(lines)
    total_pages = (total_lines + page_size - 1) // page_size
    current_page = 1
    
    # 检查是否有保存的位置
    position_file = f"{file_path}.position"
    if os.path.exists(position_file):
        with open(position_file, 'r', encoding='utf-8') as f:
            try:
                saved_page = int(f.read().strip())
                if 1 <= saved_page <= total_pages:
                    current_page = saved_page
                    print(f"已恢复到第 {current_page} 页")
            except:
                pass
    
    while True:
        # 计算当前页的起始和结束行
        start_line = (current_page - 1) * page_size
        end_line = min(current_page * page_size, total_lines)
        
        # 清屏
        os.system('cls' if os.name == 'nt' else 'clear')
        
        # 显示当前页内容
        print(f"文件: {file_path}")
        print(f"第 {current_page}/{total_pages} 页 (共 {total_lines} 行)")
        print("-" * 80)
        
        for i in range(start_line, end_line):
            print(f"{i+1:4d}: {lines[i].rstrip()}")
        
        print("-" * 80)
        print("命令: n(下一页) | p(上一页) | g(跳转页) | q(退出)")
        
        # 保存当前位置
        with open(position_file, 'w', encoding='utf-8') as f:
            f.write(str(current_page))
        
        # 等待用户输入
        while True:
            command = input("请输入命令: ").strip().lower()
            if command == 'n':
                if current_page < total_pages:
                    current_page += 1
                else:
                    print("已经是最后一页")
                    input("按回车键继续...")
                break
            elif command == 'p':
                if current_page > 1:
                    current_page -= 1
                else:
                    print("已经是第一页")
                    input("按回车键继续...")
                break
            elif command == 'g':
                try:
                    page_num = int(input("请输入页码: ").strip())
                    if 1 <= page_num <= total_pages:
                        current_page = page_num
                    else:
                        print(f"页码必须在 1 到 {total_pages} 之间")
                        input("按回车键继续...")
                except ValueError:
                    print("请输入有效的页码")
                    input("按回车键继续...")
                break
            elif command == 'q':
                print("退出查看")
                return
            else:
                print("无效命令，请重新输入")

if __name__ == '__main__':
    file_path = r'd:\learning-area-english\zdc.md'
    view_file(file_path)
