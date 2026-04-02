# 文档对比工具

支持对比 PDF、DOCX、PPTX、TXT 格式文档，生成 Word 格式的差异报告。

## 功能特性

- 支持多种格式：**PDF**、DOCX、PPTX、TXT
- PDF 智能处理：跨页段落合并、页码过滤、视觉行号提取
- 横向页面布局，仅显示差异行
- 单词级差异高亮（蓝色/红色）
- 句子级缺失检测，显示 `[此处缺失句子]`（绿色）
- 自动合并连续的缺失占位符

## 使用方法

```bash
python compare_docs.py <文件1> <文件2>
```

输出：`Comparison_文件名1_VS_文件名2.docx`

## 依赖安装

```bash
pip install python-docx pdfplumber python-pptx
```

## 作者

yuxiacn-dev
