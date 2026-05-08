#!/usr/bin/env python3
import csv
import chromadb
import os
import sys
import argparse


def main():
    """将 product_library.csv 导入 ChromaDB"""
    
    # 设置目录路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = script_dir  # CSV 文件目录
    persist_directory = os.path.join(script_dir, "..", "chroma_db")  # ChromaDB 数据存储目录
    csv_path = os.path.join(data_dir, "product_library.csv")
    collection_name = "product_library"
    
    print("=" * 60)
    print("产品库导入工具")
    print("=" * 60)
    
    # 检查 CSV 文件是否存在
    if not os.path.exists(csv_path):
        print(f"错误: CSV 文件不存在: {csv_path}")
        sys.exit(1)
    
    print(f"\nCSV 文件: {csv_path}")
    print(f"ChromaDB 路径: {persist_directory}")
    print(f"集合名称: {collection_name}")
    
    # 确保 ChromaDB 目录存在
    os.makedirs(persist_directory, exist_ok=True)
    
    # 连接 ChromaDB
    client = chromadb.PersistentClient(path=persist_directory)
    
    # 检查集合是否已存在
    existing_collections = [col.name for col in client.list_collections()]
    collection_exists = collection_name in existing_collections
    
    if collection_exists:
        collection = client.get_collection(name=collection_name)
        print(f"\n集合 '{collection_name}' 已存在，当前有 {collection.count()} 条记录")
        response = input("是否要更新集合？(y/n): ").strip().lower()
        if response != 'y':
            print("操作已取消。")
            sys.exit(0)
    else:
        print(f"\n创建新集合: {collection_name}")
        collection = client.create_collection(name=collection_name)
    
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
    if '申报要素' in fieldnames:
        vector_fields.append('申报要素')
    if '税号' in fieldnames:
        vector_fields.append('税号')
    
    print(f"向量化使用字段: {', '.join(vector_fields)}")
    
    # 准备要插入的数据
    documents = []  # 存储向量化的文本内容
    metadatas = []  # 存储元数据（所有 CSV 字段）
    ids = []  # 存储每条记录的唯一 ID
    
    for idx, row in enumerate(all_rows, 1):
        # 构建用于向量化的文档文本
        doc_parts = []
        for key in vector_fields:
            value = row.get(key, '')
            if value:
                doc_parts.append(f"{value}")
        doc_text = " ".join(doc_parts)
        
        documents.append(doc_text)
        metadatas.append({k: v for k, v in row.items()})  # 所有字段都作为元数据保存
        
        # 生成唯一 ID
        if 'ID' in row:
            row_id = row['ID']
        elif 'id' in row:
            row_id = row['id']
        else:
            row_id = str(idx)
        ids.append(f"{collection_name}_{row_id}")
    
    # 如果集合已存在，先删除旧数据
    if collection_exists:
        existing_ids = collection.get()["ids"]
        if existing_ids:
            print(f"\n删除 {len(existing_ids)} 条旧记录...")
            collection.delete(ids=existing_ids)
    
    # 添加新数据到集合
    print(f"\n正在导入 {len(documents)} 条记录...")
    collection.add(
        documents=documents,
        metadatas=metadatas,
        ids=ids
    )
    
    print(f"\n" + "=" * 60)
    print(f"成功！集合 '{collection_name}' 已更新")
    print(f"当前记录数: {collection.count()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
