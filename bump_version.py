#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
版本号自动更新工具
自动递增版本号的修订号和日期

使用方法:
    python bump_version.py [minor]
    
参数:
    minor - 可选，如果提供则递增次版本号（如 V2.0 → V2.1）

示例:
    python bump_version.py          # 仅更新日期和修订号: V2.0.Build20260403.1 → V2.0.Build20260403.2
    python bump_version.py minor    # 更新次版本号: V2.0.Build20260403.1 → V2.1.Build20260403.1
"""

import re
import sys
from datetime import datetime
from pathlib import Path


def bump_version(minor=False):
    """更新版本号"""
    
    file_path = Path(__file__).parent / "compare_docs.py"
    
    if not file_path.exists():
        print(f"错误: 找不到文件 {file_path}")
        return False
    
    content = file_path.read_text(encoding='utf-8')
    
    # 查找当前版本号
    version_pattern = r'VERSION = "V(\d+)\.(\d+)\.Build(\d{8})\.(\d+)"'
    match = re.search(version_pattern, content)
    
    if not match:
        print("错误: 找不到版本号")
        return False
    
    major = int(match.group(1))
    minor_ver = int(match.group(2))
    old_date = match.group(3)
    build = int(match.group(4))
    
    # 获取当前日期
    today = datetime.now().strftime("%Y%m%d")
    
    if minor:
        # 递增次版本号，重置修订号
        minor_ver += 1
        build = 1
        print(f"更新次版本号: V{major}.{minor_ver-1} → V{major}.{minor_ver}")
    else:
        # 检查日期是否变化
        if today == old_date:
            # 同一天，递增修订号
            build += 1
        else:
            # 日期变化，重置修订号
            build = 1
        print(f"更新修订号: Build{old_date}.{build-1 if today==old_date else '?'} → Build{today}.{build}")
    
    # 构建新版本号
    new_version = f'V{major}.{minor_ver}.Build{today}.{build}'
    
    # 替换版本号
    new_content = re.sub(version_pattern, f'VERSION = "{new_version}"', content)
    
    # 写回文件
    file_path.write_text(new_content, encoding='utf-8')
    
    print(f"版本号已更新: {new_version}")
    return True


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'minor':
        bump_version(minor=True)
    else:
        bump_version(minor=False)
