#!/usr/bin/env python3
import os
import sys
import argparse
import json
from pathlib import Path


def get_excel_sheets(file_path):
    """
    获取Excel文件的所有sheet名称
    
    优先使用openpyxl处理.xlsx，失败则尝试xlrd处理.xls
    
    Args:
        file_path: Excel文件路径
        
    Returns:
        sheet名称列表，如果读取失败返回空列表
    """
    suffix = Path(file_path).suffix.lower()
    
    # 尝试用openpyxl读取.xlsx文件
    if suffix in ['.xlsx']:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except ImportError:
            pass
        except Exception as e:
            pass
    
    # 尝试用xlrd读取.xls或.xlsx（作为备用方案）
    if suffix in ['.xls'] or True:
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet_names = wb.sheet_names()
            return sheet_names
        except ImportError:
            pass
        except Exception as e:
            pass
    
    return []


def check_files(input_path):
    """
    检查指定路径下的文件，统计文件数量并分类
    
    Args:
        input_path: 文件或目录路径
        
    Returns:
        包含文件统计信息的字典：
        {
            "total_count": 总文件数,
            "file_types": {"excel": 数量, "word": 数量, "txt": 数量, "other": 数量},
            "files": 文件信息列表
        }
    """
    file_info = {
        "total_count": 0,
        "file_types": {
            "excel": 0,
            "word": 0,
            "txt": 0,
            "other": 0
        },
        "files": []
    }
    
    input_path = Path(input_path)
    
    # 检查路径是否存在
    if not input_path.exists():
        return file_info
    
    files_to_check = []
    
    # 收集需要检查的文件列表
    if input_path.is_file():
        files_to_check.append(input_path)
    elif input_path.is_dir():
        for item in input_path.iterdir():
            if item.is_file():
                files_to_check.append(item)
    
    # 处理每个文件
    for file_idx, file_path in enumerate(files_to_check):
        file_name = file_path.name
        suffix = file_path.suffix.lower()
        
        file_type = "other"
        
        # Excel文件处理：按sheet拆分
        if suffix in ['.xlsx', '.xls']:
            file_type = "excel"
            sheet_names = get_excel_sheets(str(file_path))
            if sheet_names:
                for sheet_idx, sheet_name in enumerate(sheet_names):
                    file_item = {
                        "name": f"{file_name}#{sheet_name}",
                        "type": file_type,
                        "file_name": file_name,
                        "sheet_name": sheet_name
                    }
                    file_info["files"].append(file_item)
                    file_info["total_count"] += 1
                    file_info["file_types"][file_type] += 1
            else:
                file_item = {
                    "name": file_name,
                    "type": file_type,
                    "file_name": file_name
                }
                file_info["files"].append(file_item)
                file_info["total_count"] += 1
                file_info["file_types"][file_type] += 1
        # Word文件处理
        elif suffix in ['.docx', '.doc']:
            file_type = "word"
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
        # TXT文件处理
        elif suffix == '.txt':
            file_type = "txt"
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
        # 其他文件类型
        else:
            file_type = "other"
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
    
    return file_info


def main():
    """
    主函数：解析命令行参数并执行文件检查
    """
    parser = argparse.ArgumentParser(description="文件检查工具")
    parser.add_argument("path", help="文件或目录路径")
    parser.add_argument("--json", action="store_true", help="以JSON格式输出")

    args = parser.parse_args()
    
    print(f"开始检查路径: {args.path}")
    result = check_files(args.path)
    print(f"检查完成，共发现 {result['total_count']} 个文件/Sheet")
    
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("=" * 60)
        print("文件检查结果")
        print("=" * 60)
        print(f"总文件数: {result['total_count']}")
        print(f"Excel: {result['file_types']['excel']}")
        print(f"Word: {result['file_types']['word']}")
        print(f"Txt: {result['file_types']['txt']}")
        print(f"其他: {result['file_types']['other']}")
        print("\n文件列表:")
        for i, file_item in enumerate(result['files'], 1):
            print(f"  {i}. {file_item['name']} ({file_item['type']})")
        print("=" * 60)


if __name__ == "__main__":
    main()
