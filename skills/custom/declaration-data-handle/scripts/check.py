#!/usr/bin/env python3
import os
import sys
import argparse
import json
from pathlib import Path
from datetime import datetime

from utils.logging import log_message, set_debug_mode


def get_excel_sheets(file_path):
    """获取Excel文件的所有sheet名称 - 同时支持 .xlsx/.xls 格式"""
    log_message(f"开始读取Excel文件的sheet列表: {file_path}")
    suffix = Path(file_path).suffix.lower()
    
    # 优先处理 xlsx 使用 openpyxl
    if suffix in ['.xlsx']:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            log_message(f"使用 openpyxl 读取 Excel (.xlsx) 成功，共找到 {len(sheet_names)} 个sheet: {sheet_names}")
            return sheet_names
        except ImportError:
            log_message("警告: openpyxl 未安装，尝试使用 xlrd...", "WARNING")
        except Exception as e:
            log_message(f"openpyxl 读取失败，尝试使用 xlrd: {e}", "WARNING")
    
    # 对 .xls 或 openpyxl 失败的情况，使用 xlrd
    if suffix in ['.xls'] or True:
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet_names = wb.sheet_names()
            log_message(f"使用 xlrd 读取 Excel 成功，共找到 {len(sheet_names)} 个sheet: {sheet_names}")
            return sheet_names
        except ImportError:
            log_message("警告: xlrd 未安装，无法读取Excel文件", "WARNING")
        except Exception as e:
            log_message(f"xlrd 读取Excel文件失败: {e}", "ERROR")
    
    return []


def check_files(input_path):
    """
    检查指定路径下的文件
    
    参数:
        input_path: 文件路径（可以是单个文件或目录）
    
    返回:
        字典，包含文件数量和文件列表
    """
    log_message("=" * 60)
    log_message("文件检查工具启动")
    log_message("=" * 60)
    log_message(f"输入路径: {input_path}")
    
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
    
    if not input_path.exists():
        log_message(f"错误: 路径不存在: {input_path}", "ERROR")
        return file_info
    
    files_to_check = []
    
    if input_path.is_file():
        log_message("输入路径是单个文件")
        files_to_check.append(input_path)
    elif input_path.is_dir():
        log_message("输入路径是目录，开始扫描目录内文件")
        for item in input_path.iterdir():
            if item.is_file():
                log_message(f"发现待检查文件: {item.name}")
                files_to_check.append(item)
    
    log_message(f"待检查文件总数: {len(files_to_check)}")
    
    for file_idx, file_path in enumerate(files_to_check):
        file_name = file_path.name
        suffix = file_path.suffix.lower()
        log_message(f"处理第 {file_idx+1} 个文件: {file_name}, 后缀: {suffix}")
        
        file_type = "other"
        
        if suffix in ['.xlsx', '.xls']:
            file_type = "excel"
            log_message("识别为 Excel 文件类型")
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
                    log_message(f"Excel sheet {sheet_idx+1}: 已添加文件项 {file_item['name']}")
            else:
                file_item = {
                    "name": file_name,
                    "type": file_type,
                    "file_name": file_name
                }
                file_info["files"].append(file_item)
                file_info["total_count"] += 1
                file_info["file_types"][file_type] += 1
                log_message(f"无法读取sheet列表，直接添加Excel文件: {file_name}")
        elif suffix in ['.docx', '.doc']:
            file_type = "word"
            log_message("识别为 Word 文件类型")
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
            log_message(f"已添加 Word 文件: {file_name}")
        elif suffix == '.txt':
            file_type = "txt"
            log_message("识别为 TXT 文件类型")
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
            log_message(f"已添加 TXT 文件: {file_name}")
        else:
            file_type = "other"
            log_message(f"识别为其他文件类型: {suffix}")
            file_item = {
                "name": file_name,
                "type": file_type
            }
            file_info["files"].append(file_item)
            file_info["total_count"] += 1
            file_info["file_types"][file_type] += 1
            log_message(f"已添加其他类型文件: {file_name}")
    
    log_message("-" * 60)
    log_message(f"文件检查完成，总文件数: {file_info['total_count']}")
    log_message(f"文件类型统计: Excel={file_info['file_types']['excel']}, Word={file_info['file_types']['word']}, Txt={file_info['file_types']['txt']}, Other={file_info['file_types']['other']}")
    log_message("=" * 60)
    
    return file_info


def main():
    parser = argparse.ArgumentParser(description="文件检查工具")
    parser.add_argument("path", help="文件或目录路径")
    parser.add_argument("--json", action="store_true", help="以JSON格式输出")
    parser.add_argument("--debug", action="store_true", help="启用debug模式，实时保存日志到文件")
    
    args = parser.parse_args()
    
    # 设置debug模式
    set_debug_mode(args.debug)
    if args.debug:
        print("[DEBUG] Debug模式已启用，日志将保存到文件")
    
    log_message("开始解析命令行参数")
    log_message(f"命令行参数解析完成: path={args.path}, json_output={args.json}, debug_mode={args.debug}")
    
    result = check_files(args.path)
    
    if args.json:
        log_message("以JSON格式输出结果")
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        log_message("以人类可读格式输出结果")
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
    
    log_message("文件检查工具执行完毕")


if __name__ == "__main__":
    main()
