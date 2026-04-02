#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 Word COM 接口精确获取视觉行号
仅在 Windows + 安装 Microsoft Word 的环境下有效
"""

import os
import sys
import json
import tempfile
from pathlib import Path


def get_word_visual_line_numbers_win32(doc_path, output_json=None):
    """
    使用 Windows COM 接口获取 Word 文档的视觉行号
    需要 Windows 系统 + Microsoft Word 安装
    
    返回: [(paragraph_index, visual_line_number, text), ...]
    """
    try:
        import win32com.client
    except ImportError:
        print("错误: 需要安装 pywin32")
        print("pip install pywin32")
        return None
    
    if output_json is None:
        output_json = os.path.join(tempfile.gettempdir(), 'word_line_numbers.json')
    
    # 转换为绝对路径
    doc_path = os.path.abspath(doc_path)
    
    try:
        # 启动 Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        # 打开文档
        doc = word.Documents.Open(doc_path)
        
        # 获取行号信息
        line_numbers = []
        
        for i, para in enumerate(doc.Paragraphs, 1):
            text = para.Range.Text.strip()
            if not text:
                continue
                
            # 获取该段落的起始行号
            # 通过 Range.Information 获取页码和行号信息
            try:
                # wdActiveEndPageNumber = 3
                page_num = para.Range.Information(3)
                # wdFirstCharacterLineNumber = 10
                line_num = para.Range.Information(10)
                
                line_numbers.append({
                    'paragraph_index': i,
                    'page_number': page_num,
                    'line_number': line_num,
                    'text': text[:100]  # 前100字符用于验证
                })
            except Exception as e:
                line_numbers.append({
                    'paragraph_index': i,
                    'page_number': None,
                    'line_number': None,
                    'text': text[:100],
                    'error': str(e)
                })
        
        # 保存结果
        with open(output_json, 'w', encoding='utf-8') as f:
            json.dump(line_numbers, f, ensure_ascii=False, indent=2)
        
        # 关闭文档
        doc.Close(SaveChanges=False)
        word.Quit()
        
        return line_numbers
        
    except Exception as e:
        print(f"错误: {e}")
        return None


def get_visual_line_map(doc_path):
    """
    获取文档的视觉行号映射表
    返回: [visual_line_num, ...] 与段落一一对应
    """
    # 首先尝试使用 COM 接口（Windows + Word）
    result = get_word_visual_line_numbers_win32(doc_path)
    
    if result is not None:
        # 提取视觉行号列表
        visual_map = []
        for item in result:
            line_num = item.get('line_number')
            if line_num:
                visual_map.append(line_num)
            else:
                # 如果获取失败，使用段落序号
                visual_map.append(item['paragraph_index'])
        return visual_map
    
    return None


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python word_line_numbers.py <word文档路径>")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    result = get_word_visual_line_numbers_win32(doc_path)
    
    if result:
        print(f"找到 {len(result)} 个段落:")
        for item in result[:10]:  # 显示前10个
            print(f"  段落 {item['paragraph_index']}: 行号={item['line_number']}, 页码={item['page_number']}")
            print(f"    文本: {item['text'][:50]}...")
    else:
        print("获取失败")
