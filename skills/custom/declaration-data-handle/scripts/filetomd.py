#!/usr/bin/env python3
import logging
import sys
import argparse
import pandas as pd
from pathlib import Path

from utils.logging import setup_logging, log_message, set_debug_mode

logger = setup_logging()

try:
    from markitdown import MarkItDown
except ImportError:
    logger.error("MarkItDown 模块未安装，请先安装后运行本脚本")
    sys.exit(1)

def convert_single_file(input_path: Path, output_dir: Path, md: MarkItDown) -> list[str]:
    """Convert a single file to Markdown, return list of saved file paths"""
    saved_files = []
    logger.info(f"开始处理文件: {input_path.name}")
    
    if input_path.suffix.lower() in ['.xlsx', '.xls']:
        # Excel file: each sheet as separate MD file
        excel_file = pd.ExcelFile(input_path)
        logger.info(f"  检测到 Excel 文件，包含 {len(excel_file.sheet_names)} 个 Sheet")
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(input_path, sheet_name=sheet_name)
            md_content = f"# Sheet: {sheet_name}\n\n"
            md_content += df.to_markdown(index=False)
            
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            output_filename = f"{input_path.stem}_Sheet_{safe_sheet_name}.md"
            output_path = output_dir / output_filename
            
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(md_content, encoding='utf-8')
            saved_files.append(str(output_path.resolve()))
            logger.info(f"  [Sheet] {sheet_name} -> {output_path.name}")
    else:
        # Other file types
        result = md.convert(str(input_path))
        output_path = output_dir / f"{input_path.stem}.md"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(result.text_content, encoding='utf-8')
        saved_files.append(str(output_path.resolve()))
        logger.info(f"  [File] {input_path.name} -> {output_path.name}")
    
    return saved_files


def convert_directory(input_dir: Path, output_dir: Path) -> dict:
    """Convert all supported files in a directory"""
    md = MarkItDown()
    all_saved = []
    
    supported_exts = ['.pdf', '.docx', '.xlsx', '.xls', '.pptx', '.jpg', '.jpeg', '.png', '.zip', '.txt', '.html']
    logger.info(f"开始批量转换目录: {input_dir.resolve()}")
    
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for file_path in input_dir.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in supported_exts:
            saved = convert_single_file(file_path, output_dir, md)
            all_saved.extend(saved)
    
    return {
        "total_converted": len(all_saved),
        "files": all_saved,
        "output_dir": str(output_dir.resolve())
    }


def main():
    parser = argparse.ArgumentParser(description='Convert documents to Markdown (single file or directory batch)')
    parser.add_argument('input_path', help='Path to input file or directory')
    parser.add_argument('-o', '--output', help='Path to output Markdown file or output directory')
    parser.add_argument('--debug', action='store_true', help='启用debug模式，实时保存日志到文件')
    
    args = parser.parse_args()
    
    set_debug_mode(args.debug)
    if args.debug:
        print("[DEBUG] Debug模式已启用，日志将保存到文件")
    
    input_path = Path(args.input_path)
    logger.info(f"收到输入路径: {input_path.resolve()}")
    
    if not input_path.exists():
        logger.error(f"错误：输入路径不存在: {input_path}")
        sys.exit(1)
    
    md = MarkItDown()
    
    if input_path.is_file():
        # Single file conversion
        logger.info("模式：单个文件转换")
        if args.output:
            output_path = Path(args.output)
            output_dir = output_path.parent
            output_dir.mkdir(parents=True, exist_ok=True)
        else:
            output_dir = input_path.parent
            output_path = input_path.with_suffix('.md')
        
        saved = convert_single_file(input_path, output_dir, md)
        logger.info(f"完成！已转换 {len(saved)} 个文件")
        return saved[0] if saved else None
    else:
        # Directory batch conversion
        logger.info("模式：目录批量转换")
        output_dir = Path(args.output) if args.output else input_path / "markdown"
        result = convert_directory(input_path, output_dir)
        logger.info(f"完成！已转换 {result['total_converted']} 个 Markdown 文件到: {result['output_dir']}")
        return result['output_dir']


if __name__ == '__main__':
    main()
