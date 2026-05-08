#!/usr/bin/env python3
import csv
import os
import sys
import argparse
import shutil
import lancedb
import numpy as np
from sentence_transformers import SentenceTransformer


def main():
    """将 product_library.csv 导入 LanceDB"""
    
    # 设置目录路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = script_dir  # CSV 文件目录
    persist_directory = os.path.join(script_dir, "..", "lance_db")  # LanceDB 数据存储目录
    csv_path = os.path.join(data_dir, "product_library.csv")
    table_name = "product_library"
    
    print("=" * 60)
    print("产品库导入工具 (LanceDB)")
    print("=" * 60)
    
    # 检查 CSV 文件是否存在
    if not os.path.exists(csv_path):
        print(f"错误: CSV 文件不存在: {csv_path}")
        sys.exit(1)
    
    print(f"\nCSV 文件: {csv_path}")
    print(f"LanceDB 路径: {persist_directory}")
    print(f"表名称: {table_name}")
    
    # 确保 LanceDB 目录存在
    os.makedirs(persist_directory, exist_ok=True)
    
    # 连接 LanceDB
    db = lancedb.connect(persist_directory)
    
    # 检查表是否已存在，如果存在则删除
    if table_name in db.list_tables():
        table = db.open_table(table_name)
        count = len(table)
        print(f"\n表 '{table_name}' 已存在（{count} 条记录），正在删除...")
        db.drop_table(table_name)
        # 强制删除表目录（确保文件系统层面也删除）
        table_dir = os.path.join(persist_directory, f"{table_name}.lance")
        if os.path.exists(table_dir):
            shutil.rmtree(table_dir)
            print(f"✓ 表目录已删除: {table_dir}")
        print(f"✓ 表 '{table_name}' 已删除")
        # 重新连接数据库
        db = lancedb.connect(persist_directory)
    
    # 读取 CSV
    print(f"\n读取 CSV 文件...")
    all_rows = []
    fieldnames = []
    
    try:
        with open(csv_path, mode='r', encoding='utf-8-sig') as file:
            csv_reader = csv.DictReader(file)
            fieldnames = csv_reader.fieldnames
            for row in csv_reader:
                all_rows.append(row)
        
        print(f"检测到的列: {', '.join(fieldnames)}")
        print(f"共 {len(all_rows)} 条记录")
        
    except Exception as e:
        print(f"读取 CSV 失败: {e}")
        sys.exit(1)
    
    # 选择用于向量化的字段（优先使用英文描述）
    print("\n正在准备数据...")
    
    # 推荐用于向量化的字段
    vector_fields = []
    if '英文描述' in fieldnames:
        vector_fields.append('英文描述')
    if '品名' in fieldnames:
        vector_fields.append('品名')
    
    print(f"向量化使用字段: {', '.join(vector_fields)}")
    print("\n正在加载嵌入模型 (all-MiniLM-L6-v2)...")
    
    # 使用与ChromaDB相同的嵌入模型
    model = SentenceTransformer('all-MiniLM-L6-v2')
    
    # 准备要插入的数据
    data = []
    documents = []
    
    for idx, row in enumerate(all_rows, 1):
        # 构建用于向量化的文档文本
        doc_parts = []
        for key in vector_fields:
            value = row.get(key, '')
            if value:
                doc_parts.append(f"{value}")
        doc_text = " ".join(doc_parts)
        documents.append(doc_text)
        
        # 生成唯一 ID
        if 'ID' in row:
            row_id = row['ID']
        elif 'id' in row:
            row_id = row['id']
        else:
            row_id = str(idx)
        
        item = {
            "id": f"{table_name}_{row_id}",
            "document": doc_text,
        }
        # 添加所有字段都作为字段保存
        item.update(row)
        data.append(item)
    
    # 生成向量
    print("\n正在生成向量嵌入...")
    embeddings = model.encode(documents, show_progress_bar=True)
    
    # 添加向量到数据
    for i, item in enumerate(data):
        item["vector"] = embeddings[i].tolist()
    
    # 确保表不存在，如果存在则再次删除
    if table_name in db.list_tables():
        print("警告: 表仍然存在，再次删除...")
        db.drop_table(table_name)
        # 强制删除表目录
        table_dir = os.path.join(persist_directory, f"{table_name}.lance")
        if os.path.exists(table_dir):
            shutil.rmtree(table_dir)
        db = lancedb.connect(persist_directory)
    
    # 创建新表并添加数据
    print(f"\n正在导入 {len(data)} 条记录...")
    table = db.create_table(table_name, data=data)
    
    print(f"\n" + "=" * 60)
    print(f"成功！表 '{table_name}' 已更新")
    print(f"当前记录数: {len(table)}")
    print("=" * 60)


if __name__ == "__main__":
    main()

