#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用 LibreOffice 获取 Word 文档的视觉行号
跨平台方案（支持 Mac/Linux/Windows）
需要先安装 LibreOffice
"""

import os
import sys
import json
import tempfile
import subprocess
from pathlib import Path


def get_line_numbers_with_libreoffice(doc_path, output_dir=None):
    """
    使用 LibreOffice 将文档转换为 PDF，然后分析页面布局获取行号
    这是一个近似方案，可以获得较好的估算
    """
    if output_dir is None:
        output_dir = tempfile.gettempdir()
    
    doc_path = Path(doc_path).resolve()
    output_dir = Path(output_dir)
    
    # 转换为 PDF
    pdf_path = output_dir / f"{doc_path.stem}_lines.pdf"
    
    try:
        # LibreOffice 命令行转换
        cmd = [
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            str(doc_path)
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        
        if result.returncode != 0:
            print(f"LibreOffice 转换失败: {result.stderr}")
            return None
        
        # 生成的 PDF 路径
        generated_pdf = output_dir / f"{doc_path.stem}.pdf"
        if not generated_pdf.exists():
            print(f"PDF 文件未生成")
            return None
        
        # 使用 pdfplumber 分析 PDF 获取行信息
        try:
            import pdfplumber
        except ImportError:
            print("需要安装 pdfplumber: pip install pdfplumber")
            return None
        
        line_numbers = []
        line_counter = 1
        
        with pdfplumber.open(str(generated_pdf)) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # 获取页面文本，保留布局信息
                words = page.extract_words()
                
                # 按 y 坐标分组（同一行的词）
                lines_dict = {}
                for word in words:
                    y = round(float(word['top']), 1)  # 取一位小数分组
                    if y not in lines_dict:
                        lines_dict[y] = []
                    lines_dict[y].append(word)
                
                # 按 y 坐标排序，构建行信息
                for y in sorted(lines_dict.keys()):
                    words_in_line = lines_dict[y]
                    text = ' '.join(w['text'] for w in sorted(words_in_line, key=lambda x: float(x['x0'])))
                    
                    line_numbers.append({
                        'page': page_num,
                        'line_in_page': len([l for l in line_numbers if l['page'] == page_num]) + 1,
                        'global_line': line_counter,
                        'text': text[:100],
                        'y_position': y
                    })
                    line_counter += 1
        
        # 清理临时 PDF
        try:
            generated_pdf.unlink()
        except:
            pass
        
        return line_numbers
        
    except subprocess.TimeoutExpired:
        print("LibreOffice 转换超时")
        return None
    except FileNotFoundError:
        print("错误: 未找到 LibreOffice (soffice)")
        print("请先安装 LibreOffice:")
        print("  Mac: brew install --cask libreoffice")
        print("  Ubuntu: sudo apt install libreoffice")
        print("  Windows: 官网下载安装")
        return None
    except Exception as e:
        print(f"错误: {e}")
        return None


def estimate_line_numbers_precise(doc_path, chars_per_line=80):
    """
    高精度估算方案：结合段落样式和字符宽度
    适用于无法使用外部工具的情况
    """
    from docx import Document
    from docx.shared import Inches, Pt
    
    doc = Document(doc_path)
    
    # 获取页面设置
    section = doc.sections[0]
    page_width = section.page_width.inches if section.page_width else 8.5
    left_margin = section.left_margin.inches if section.left_margin else 1.0
    right_margin = section.right_margin.inches if section.right_margin else 1.0
    
    # 可用宽度（英寸）
    available_width = page_width - left_margin - right_margin
    
    # 默认字符宽度（英寸）
    # 12pt 字体，约每英寸 10 个字符
    default_char_width = 0.1  # 英寸
    
    line_numbers = []
    current_line = 1
    
    for i, para in enumerate(doc.paragraphs, 1):
        text = para.text.strip()
        if not text:
            continue
        
        # 尝试获取段落字体大小
        font_size = 12  # 默认 12pt
        try:
            if para.runs:
                font = para.runs[0].font
                if font.size:
                    font_size = font.size.pt
        except:
            pass
        
        # 计算该字体下的字符宽度
        # 粗略估计：字体越大，字符越宽
        char_width = default_char_width * (font_size / 12)
        
        # 计算每行可容纳的字符数
        chars_per_line_actual = int(available_width / char_width)
        chars_per_line_actual = max(40, min(chars_per_line_actual, 100))  # 限制在 40-100 之间
        
        # 考虑缩进
        indent_chars = 0
        if para.paragraph_format.left_indent:
            indent_inches = para.paragraph_format.left_indent.inches
            indent_chars = int(indent_inches / char_width)
        
        # 计算需要的行数
        available_chars_per_line = chars_per_line_actual - indent_chars
        text_length = len(text)
        
        # 中英文混合计算（中文算 2 个字符宽度）
        effective_length = 0
        for char in text:
            if '\u4e00' <= char <= '\u9fff':
                effective_length += 2
            else:
                effective_length += 1
        
        lines_needed = max(1, (effective_length + available_chars_per_line - 1) // available_chars_per_line)
        
        line_numbers.append({
            'paragraph_index': i,
            'start_line': current_line,
            'end_line': current_line + lines_needed - 1,
            'lines_count': lines_needed,
            'text': text[:80],
            'font_size': font_size,
            'chars_per_line': available_chars_per_line
        })
        
        current_line += lines_needed
    
    return line_numbers


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python word_line_numbers_libreoffice.py <word文档路径>")
        print("\n说明:")
        print("1. 首先尝试使用 LibreOffice (最精确)")
        print("2. 如果失败，使用高精度估算")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    
    # 首先尝试 LibreOffice
    print("尝试使用 LibreOffice...")
    result = get_line_numbers_with_libreoffice(doc_path)
    
    if result:
        print(f"LibreOffice 成功！找到 {len(result)} 行")
        for item in result[:5]:
            print(f"  第 {item['global_line']} 行 (第 {item['page']} 页): {item['text'][:50]}...")
    else:
        print("LibreOffice 失败，使用高精度估算...")
        result = estimate_line_numbers_precise(doc_path)
        
        if result:
            print(f"估算完成！共 {len(result)} 个段落")
            for item in result[:5]:
                print(f"  段落 {item['paragraph_index']}: 起始行 {item['start_line']}, 占用 {item['lines_count']} 行")
