#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Version Auto-Update Tool
Automatically increment version number revision and date

Usage:
    python bump_version.py [filename] [minor]
    
Arguments:
    filename - Optional, target file to update (default: auto-detect)
    minor    - Optional, if provided increment minor version (e.g., V2.0 → V2.1)

Examples:
    python bump_version.py                     # Auto-detect file, update revision
    python bump_version.py minor               # Auto-detect file, update minor version
    python bump_version.py comparedocen.py     # Update specific file revision
    python bump_version.py comparedocen.py minor  # Update specific file minor version
"""

import re
import sys
from datetime import datetime
from pathlib import Path


def bump_version(target_file=None, minor=False):
    """Update version number"""
    
    # If target file specified, use it; otherwise auto-detect
    if target_file:
        file_path = Path(__file__).parent / target_file
        if not file_path.exists():
            print(f"Error: File not found: {target_file}")
            return False
    else:
        # Auto-detect from possible files
        possible_files = ['comparedocscn.py', 'comparedocen.py', 'compare_docs.py']
        file_path = None
        for fname in possible_files:
            fpath = Path(__file__).parent / fname
            if fpath.exists():
                file_path = fpath
                break
        
        if file_path is None:
            print(f"Error: No file found, tried: {', '.join(possible_files)}")
            return False
    
    content = file_path.read_text(encoding='utf-8')
    
    # Find current version
    version_pattern = r'VERSION = "V(\d+)\.(\d+) Build(\d{8})\.(\d+)"'
    match = re.search(version_pattern, content)
    
    if not match:
        print("Error: Version number not found")
        return False
    
    major = int(match.group(1))
    minor_ver = int(match.group(2))
    old_date = match.group(3)
    build = int(match.group(4))
    
    # Get current date
    today = datetime.now().strftime("%Y%m%d")
    
    if minor:
        # Increment minor version, reset revision
        minor_ver += 1
        build = 1
        print(f"Update minor version: V{major}.{minor_ver-1} → V{major}.{minor_ver}")
    else:
        # Check if date changed
        if today == old_date:
            # Same day, increment revision
            build += 1
        else:
            # Date changed, reset revision
            build = 1
        print(f"Update revision: Build{old_date}.{build-1 if today==old_date else '?'} → Build{today}.{build}")
    
    # Build new version string
    new_version = f'V{major}.{minor_ver} Build{today}.{build}'
    
    # Replace version
    new_content = re.sub(version_pattern, f'VERSION = "{new_version}"', content)
    
    # Write back
    file_path.write_text(new_content, encoding='utf-8')
    
    print(f"Version updated in {file_path.name}: {new_version}")
    return True


if __name__ == '__main__':
    args = sys.argv[1:]
    
    target_file = None
    minor_flag = False
    
    for arg in args:
        if arg == 'minor':
            minor_flag = True
        elif arg.endswith('.py'):
            target_file = arg
        else:
            print(f"Unknown argument: {arg}")
            print("Usage: python bump_version.py [filename.py] [minor]")
            sys.exit(1)
    
    success = bump_version(target_file=target_file, minor=minor_flag)
    sys.exit(0 if success else 1)
