#!/usr/bin/env python3
import sys
import argparse
import pandas as pd
from pathlib import Path


try:
    from markitdown import MarkItDown
except ImportError:
    print("MarkItDown 模块未安装，请先安装后运行本脚本")
    sys.exit(1)


def convert_single_file(input_path: Path, output_dir: Path, md: MarkItDown) -> list[str]:
    """
    将单个文件转换为Markdown格式
    
    对于Excel文件，每个Sheet会单独生成一个Markdown文件
    对于其他文件，使用MarkItDown进行转换
    
    Args:
        input_path: 输入文件路径
        output_dir: 输出目录路径
        md: MarkItDown实例
        
    Returns:
        保存的文件路径列表
    """
    saved_files = []
    
    # Excel文件特殊处理：每个Sheet单独转换
    if input_path.suffix.lower() in ['.xlsx', '.xls']:
        print(f"处理Excel文件: {input_path.name}")
        excel_file = pd.ExcelFile(input_path)
        for sheet_idx, sheet_name in enumerate(excel_file.sheet_names):
            print(f"  转换Sheet {sheet_idx + 1}/{len(excel_file.sheet_names)}: {sheet_name}")
            df = pd.read_excel(input_path, sheet_name=sheet_name)
            md_content = f"# Sheet: {sheet_name}\n\n"
            md_content += df.to_markdown(index=False)
            
            # 生成安全的文件名（移除特殊字符）
            safe_sheet_name = "".join(c for c in sheet_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            output_filename = f"{input_path.stem}_Sheet_{safe_sheet_name}.md"
            output_path = output_dir / output_filename
            
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(md_content, encoding='utf-8')
            saved_files.append(str(output_path.resolve()))
            print(f"  已保存: {output_filename}")
    else:
        # 其他文件使用MarkItDown转换
        print(f"处理文件: {input_path.name}")
        result = md.convert(str(input_path))
        output_path = output_dir / f"{input_path.stem}.md"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(result.text_content, encoding='utf-8')
        saved_files.append(str(output_path.resolve()))
        print(f"  已保存: {output_path.name}")
    
    return saved_files


def convert_directory(input_dir: Path, output_dir: Path) -> dict:
    """
    批量转换目录下的所有支持的文件
    
    Args:
        input_dir: 输入目录路径
        output_dir: 输出目录路径
        
    Returns:
        包含转换结果的字典：
        {
            "total_converted": 转换总数,
            "files": 保存的文件路径列表,
            "output_dir": 输出目录
        }
    """
    md = MarkItDown()
    all_saved = []
    
    # 支持的文件扩展名
    supported_exts = ['.pdf', '.docx', '.xlsx', '.xls', '.pptx', '.jpg', '.jpeg', '.png', '.zip', '.txt', '.html']
    
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"开始批量转换目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    
    # 遍历目录下的所有文件
    file_count = 0
    for file_path in input_dir.iterdir():
        if file_path.is_file() and file_path.suffix.lower() in supported_exts:
            file_count += 1
    
    print(f"找到 {file_count} 个支持的文件")
    print("-" * 60)
    
    # 逐个转换文件
    for idx, file_path in enumerate(input_dir.iterdir(), 1):
        if file_path.is_file() and file_path.suffix.lower() in supported_exts:
            print(f"[{idx}/{file_count}] 处理: {file_path.name}")
            saved = convert_single_file(file_path, output_dir, md)
            all_saved.extend(saved)
    
    print("-" * 60)
    print(f"转换完成！共转换 {len(all_saved)} 个文件")
    
    return {
        "total_converted": len(all_saved),
        "files": all_saved,
        "output_dir": str(output_dir.resolve())
    }


def main():
    """
    主函数：解析命令行参数并执行转换
    """
    parser = argparse.ArgumentParser(description='Convert documents to Markdown (single file or directory batch)')
    parser.add_argument('input_path', help='Path to input file or directory')
    parser.add_argument('-o', '--output', help='Path to output Markdown file or output directory')
    
    args = parser.parse_args()
    
    input_path = Path(args.input_path)
    
    if not input_path.exists():
        print(f"错误：输入路径不存在: {input_path}")
        sys.exit(1)
    
    md = MarkItDown()
    
    # 单个文件转换
    if input_path.is_file():
        if args.output:
            output_path = Path(args.output)
            output_dir = output_path.parent
            output_dir.mkdir(parents=True, exist_ok=True)
        else:
            output_dir = input_path.parent
            output_path = input_path.with_suffix('.md')
        
        saved = convert_single_file(input_path, output_dir, md)
        if saved:
            print(f"转换成功！输出文件: {saved[0]}")
            return saved[0]
        else:
            return None
    # 目录批量转换
    else:
        output_dir = Path(args.output) if args.output else input_path / "markdown"
        result = convert_directory(input_path, output_dir)
        return result['output_dir']


if __name__ == '__main__':
    main()
