#!/usr/bin/env python3
import os
import sys
import json
import argparse
import uuid
import re
from pathlib import Path
from datetime import datetime
from collections import defaultdict
from fuzzywuzzy import fuzz

# 尝试导入 Excel 处理库
try:
    import openpyxl
    from openpyxl.styles import Font
    from openpyxl.worksheet.datavalidation import DataValidation
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# 尝试导入向量数据库
try:
    import lancedb
    from sentence_transformers import SentenceTransformer
    LANCEDB_AVAILABLE = True
    # 预加载模型
    _embedding_model = None
except ImportError:
    LANCEDB_AVAILABLE = False

# 预编译正则表达式，提高性能
PLACEHOLDER_PATTERN = re.compile(r'\{([^}]+)\}')
MULTI_PIPE_PATTERN = re.compile(r'\|+')
PIPE_START_PATTERN = re.compile(r'^\|+')
PIPE_END_PATTERN = re.compile(r'\|+$')
MULTI_SLASH_PATTERN = re.compile(r'/+')

# 字段名到完整路径的映射
FIELD_PATH_MAP = {
    "IEFlag": "DecMessage.DecHead.IEFlag",
    "BillNo": "DecMessage.DecHead.BillNo",
    "ContrNo": "DecMessage.DecHead.ContrNo",
    "CutMode": "DecMessage.DecHead.CutMode",
    "DistinatePort": "DecMessage.DecHead.DistinatePort",
    "FeeCurr": "DecMessage.DecHead.FeeCurr",
    "FeeMark": "DecMessage.DecHead.FeeMark",
    "FeeRate": "DecMessage.DecHead.FeeRate",
    "GrossWet": "DecMessage.DecHead.GrossWet",
    "InsurCurr": "DecMessage.DecHead.InsurCurr",
    "InsurMark": "DecMessage.DecHead.InsurMark",
    "InsurRate": "DecMessage.DecHead.InsurRate",
    "ManualNo": "DecMessage.DecHead.ManualNo",
    "NetWt": "DecMessage.DecHead.NetWt",
    "NoteS": "DecMessage.DecHead.NoteS",
    "OtherCurr": "DecMessage.DecHead.OtherCurr",
    "OtherMark": "DecMessage.DecHead.OtherMark",
    "OtherRate": "DecMessage.DecHead.OtherRate",
    "OwnerCode": "DecMessage.DecHead.OwnerCode",
    "OwnerCiqCode": "DecMessage.DecHead.OwnerCiqCode",
    "OwnerName": "DecMessage.DecHead.OwnerName",
    "PackNo": "DecMessage.DecHead.PackNo",
    "TradeCountry": "DecMessage.DecHead.TradeCountry",
    "TradeMode": "DecMessage.DecHead.TradeMode",
    "TrafMode": "DecMessage.DecHead.TrafMode",
    "TrafName": "DecMessage.DecHead.TrafName",
    "VoyNo": "DecMessage.DecHead.VoyNo",
    "TransMode": "DecMessage.DecHead.TransMode",
    "WrapType": "DecMessage.DecHead.WrapType",
    "TradeArea": "DecMessage.DecHead.TradeArea",
    "DespPort": "DecMessage.DecHead.DespPort",
    "OverseasConsignorAEOCode": "DecMessage.DecHead.OverseasConsignorAEOCode",
    "OverseasConsigneeEname": "DecMessage.DecHead.OverseasConsigneeEname",
    "TradeCode": "DecMessage.DecHead.TradeCode",
    "TradeCiqCode": "DecMessage.DecHead.TradeCiqCode",
    "TradeName": "DecMessage.DecHead.TradeName",
    "TradeNameEn": "DecMessage.DecHead.TradeNameEn"
}


def reorder_results_with_keyword_match(results, search_text, text_field_names):
    """
    对结果进行重新排序，将关键词匹配度高的结果排在前面
    
    参数:
        results: 查询结果字典
        search_text: 搜索文本
        text_field_names: 元数据中用于匹配的字段名列表（必须提供）
    
    返回:
        重新排序后的结果字典
    """
    if not search_text or not results.get('ids') or len(results['ids'][0]) == 0:
        return results
    
    ids = results['ids'][0]
    docs = results['documents'][0]
    metas = results['metadatas'][0] if results.get('metadatas') else [None] * len(ids)
    distances = results['distances'][0] if results.get('distances') else [None] * len(ids)
    
    # 将搜索词拆分为关键词
    search_terms = search_text.lower().split()
    
    def count_matching_keywords(text):
        text_lower = text.lower()
        return sum(1 for term in search_terms if term in text_lower)
    
    # 为每个结果评分
    scored_results = []
    for i in range(len(ids)):
        meta = metas[i]
        doc = docs[i]
        
        # 从元数据中收集文本
        meta_texts = []
        if meta:
            for field_name in text_field_names:
                field_value = meta.get(field_name, '')
                if field_value:
                    meta_texts.append(str(field_value))
        meta_text = ' '.join(meta_texts)
        
        # 计算匹配分数
        score = 0
        search_lower = search_text.lower()
        meta_lower = meta_text.lower()
        doc_lower = doc.lower()
        
        # 1. 完全匹配（最高优先级）
        if meta_lower == search_lower:
            score = 1000
        # 2. 元数据包含整个搜索词
        elif search_lower in meta_lower:
            score = 500
        # 3. 文档包含整个搜索词
        elif search_lower in doc_lower:
            score = 200
        # 4. 计算关键词匹配数量
        else:
            keyword_count = count_matching_keywords(meta_text + ' ' + doc)
            score = keyword_count * 10
        
        scored_results.append((-score, distances[i] if distances[i] is not None else float('inf'), ids[i], docs[i], metas[i], distances[i]))
    
    # 按分数降序、距离升序排序
    scored_results.sort()
    
    # 重新构建 results 字典
    new_ids = [x[2] for x in scored_results]
    new_docs = [x[3] for x in scored_results]
    new_metas = [x[4] for x in scored_results]
    new_distances = [x[5] for x in scored_results]
    
    return {
        'ids': [new_ids],
        'documents': [new_docs],
        'metadatas': [new_metas] if new_metas[0] is not None else None,
        'distances': [new_distances] if new_distances[0] is not None else None
    }


def get_embedding_model():
    """获取或初始化嵌入模型"""
    global _embedding_model
    if _embedding_model is None:
        _embedding_model = SentenceTransformer('all-MiniLM-L6-v2')
    return _embedding_model


def query_from_db(persist_directory, collection_name, query_text=None, filter_conditions=None, 
                  n_results=10, strict_match_fields=None):
    """
    通用的向量数据库查询函数
    
    参数:
        persist_directory: LanceDB 持久化目录路径
        collection_name: 表名称
        query_text: 查询文本（用于向量搜索）
        filter_conditions: 过滤条件字典，格式为 {field_name: value}
        n_results: 返回结果数量（默认10）
        strict_match_fields: 需要严格匹配的字段名列表
    
    返回:
        查询结果字典，格式为 {'ids': [...], 'documents': [...], 'metadatas': [...], 'distances': [...]}
    """
    db = lancedb.connect(persist_directory)
    
    try:
        table = db.open_table(collection_name)
    except Exception as e:
        print(f"Error: Table '{collection_name}' not found.")
        return None
    
    # 如果有过滤条件，先进行过滤
    if filter_conditions and len(filter_conditions) > 0:
        # 获取所有数据进行过滤
        all_items = table.to_pandas()
        filtered_items = []
        
        for _, row in all_items.iterrows():
            match = True
            for key, value in filter_conditions.items():
                meta_value = str(row.get(key, ''))
                
                # 严格匹配字段：双向包含检查
                if strict_match_fields and key in strict_match_fields:
                    if not meta_value or (str(value).lower() not in meta_value.lower() and meta_value.lower() not in str(value).lower()):
                        match = False
                        break
            
            if match:
                filtered_items.append(row)
        
        if filtered_items:
            # 如果有查询文本，进行向量搜索
            if query_text:
                model = get_embedding_model()
                query_vector = model.encode(query_text).tolist()
                
                # 使用LanceDB的向量搜索
                results_df = table.search(query_vector).limit(n_results).to_pandas()
                
                # 应用过滤条件
                if strict_match_fields and filter_conditions:
                    mask = []
                    for _, row in results_df.iterrows():
                        item_match = True
                        for key, value in filter_conditions.items():
                            if key in strict_match_fields:
                                meta_value = str(row.get(key, ''))
                                if not meta_value or (str(value).lower() not in meta_value.lower() and meta_value.lower() not in str(value).lower()):
                                    item_match = False
                                    break
                        mask.append(item_match)
                    results_df = results_df[mask]
                
                # 转换为ChromaDB兼容格式
                ids = results_df['id'].tolist()
                documents = results_df['document'].tolist()
                # 将所有其他列作为metadata
                metadatas = []
                for _, row in results_df.iterrows():
                    meta = row.drop(['id', 'document', 'vector']).to_dict()
                    metadatas.append(meta)
                distances = results_df['_distance'].tolist() if '_distance' in results_df.columns else None
                
                results = {
                    'ids': [ids],
                    'documents': [documents],
                    'metadatas': [metadatas],
                    'distances': [distances] if distances else None
                }
            else:
                # 没有查询文本，直接返回所有过滤结果
                ids = [item['id'] for item in filtered_items]
                documents = [item['document'] for item in filtered_items]
                metadatas = []
                for item in filtered_items:
                    meta = {k: v for k, v in item.items() if k not in ['id', 'document', 'vector']}
                    metadatas.append(meta)
                
                results = {
                    'ids': [ids],
                    'documents': [documents],
                    'metadatas': [metadatas],
                    'distances': None
                }
        else:
            print("  No items found matching filter criteria. Searching all items...")
            if query_text:
                model = get_embedding_model()
                query_vector = model.encode(query_text).tolist()
                results_df = table.search(query_vector).limit(n_results).to_pandas()
                
                ids = results_df['id'].tolist()
                documents = results_df['document'].tolist()
                metadatas = []
                for _, row in results_df.iterrows():
                    meta = row.drop(['id', 'document', 'vector']).to_dict()
                    metadatas.append(meta)
                distances = results_df['_distance'].tolist() if '_distance' in results_df.columns else None
                
                results = {
                    'ids': [ids],
                    'documents': [documents],
                    'metadatas': [metadatas],
                    'distances': [distances] if distances else None
                }
            else:
                results = None
    else:
        # 没有过滤条件，直接查询
        if query_text:
            model = get_embedding_model()
            query_vector = model.encode(query_text).tolist()
            results_df = table.search(query_vector).limit(n_results).to_pandas()
            
            ids = results_df['id'].tolist()
            documents = results_df['document'].tolist()
            metadatas = []
            for _, row in results_df.iterrows():
                meta = row.drop(['id', 'document', 'vector']).to_dict()
                metadatas.append(meta)
            distances = results_df['_distance'].tolist() if '_distance' in results_df.columns else None
            
            results = {
                'ids': [ids],
                'documents': [documents],
                'metadatas': [metadatas],
                'distances': [distances] if distances else None
            }
        else:
            results = None
    
    return results


def log_message(message, level="INFO"):
    """记录日志信息"""
    try:
        import pandas as pd
        timestamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    except ImportError:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    output = f"[{timestamp}] [{level}] {message}"
    print(output)


def load_analysis_results(analysis_file=None, analysis_json=None):
    """加载分析结果"""
    if analysis_json:
        # 从 JSON 字符串加载
        try:
            return json.loads(analysis_json)
        except json.JSONDecodeError as e:
            log_message(f"解析 JSON 字符串失败: {e}", "ERROR")
            return None
    elif analysis_file and os.path.exists(analysis_file):
        # 从文件加载
        try:
            with open(analysis_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            log_message(f"读取分析结果文件失败: {e}", "ERROR")
            return None
    else:
        log_message("没有提供分析结果", "ERROR")
        return None


def merge_json_objects_with_guid(json_list_with_guid):
    """合并多个带 GUID 的 JSON 对象"""
    if not json_list_with_guid:
        return {}, {}

    # 第一步：先区分发票、企业默认值和其他单证
    invoice_data = None
    default_data = None
    other_data_list = []

    for data_with_guid in json_list_with_guid:
        json_data = data_with_guid.get('analysis', {}) or data_with_guid
        file_guid = data_with_guid.get('guid', str(uuid.uuid4()))
        file_business_type = json_data.get('FileBusinessType', '')

        if file_business_type == '发票':
            invoice_data = {'json': json_data, 'guid': file_guid}
        elif file_business_type == '企业默认值':
            default_data = {'json': json_data, 'guid': file_guid}
        else:
            other_data_list.append({'json': json_data, 'guid': file_guid})

    # 第二步：收集所有 DecList 数据
    dec_list_groups = {}
    existing_items = []
    group_invoice_gno = {}

    # 先处理发票，建立基准
    if invoice_data:
        json_data = invoice_data['json']
        file_guid = invoice_data['guid']

        if 'DecMessage' in json_data and 'DecLists' in json_data['DecMessage']:
            dec_list = json_data['DecMessage']['DecLists'].get('DecList', [])
            if isinstance(dec_list, list):
                for item in dec_list:
                    if isinstance(item, dict):
                        description = item.get('Description', '')
                        qty1 = item.get('Qty1')
                        gno = item.get('GNo')
                        if description:
                            qty1_str = str(qty1) if qty1 is not None else 'None'
                            group_key = f"{description}|{qty1_str}"
                            existing_items.append({
                                'desc': description,
                                'qty1': qty1,
                                'key': group_key
                            })
                            group_invoice_gno[group_key] = gno if gno is not None else float('inf')

                            if group_key not in dec_list_groups:
                                dec_list_groups[group_key] = []
                            dec_list_groups[group_key].append({
                                'guid': file_guid,
                                'item': item
                            })

    # 再处理其他单证，只根据 Description 匹配
    for data_with_guid in other_data_list:
        json_data = data_with_guid['json']
        file_guid = data_with_guid['guid']

        if 'DecMessage' in json_data and 'DecLists' in json_data['DecMessage']:
            dec_list = json_data['DecMessage']['DecLists'].get('DecList', [])
            if isinstance(dec_list, list):
                for item in dec_list:
                    if isinstance(item, dict):
                        description = item.get('Description', '')
                        if description:
                            # 查找匹配项
                            best_match, score = find_best_match_by_description(
                                description, existing_items, threshold=70
                            )

                            if best_match:
                                group_key = best_match
                                log_message(f"匹配成功: '{description}' -> {group_key} (相似度: {score}%)")
                            elif not invoice_data:
                                # 没有发票作为基准时，创建新分组
                                qty1 = item.get('Qty1')
                                qty1_str = str(qty1) if qty1 is not None else 'None'
                                group_key = f"{description}|{qty1_str}"
                                existing_items.append({
                                    'desc': description,
                                    'qty1': qty1,
                                    'key': group_key
                                })
                            else:
                                log_message(f"未找到匹配，跳过: '{description}'")
                                continue

                            if group_key not in dec_list_groups:
                                dec_list_groups[group_key] = []
                            dec_list_groups[group_key].append({
                                'guid': file_guid,
                                'item': item
                            })

    # 第三步：普通合并逻辑（不处理 DecLists）
    result = {}
    guid_to_business_type = {}

    def merge_recursive(target, source, source_guid):
        if isinstance(source, dict):
            for key, value in source.items():
                # 跳过 DecLists，单独处理
                if key == 'DecLists':
                    continue

                # 记录业务类型
                if key == 'FileBusinessType' and source_guid:
                    guid_to_business_type[source_guid] = value
                    # 只保存第一个非空值
                    if key not in target and value and value != '':
                        target[key] = value
                    continue

                if key not in target:
                    target[key] = {} if isinstance(value, (dict, list)) else []

                if isinstance(value, dict):
                    if not isinstance(target[key], dict):
                        target[key] = {}
                    merge_recursive(target[key], value, source_guid)
                elif isinstance(value, list):
                    if not isinstance(target[key], list):
                        target[key] = []
                    for list_item in value:
                        if isinstance(list_item, dict):
                            found = False
                            for existing_item in target[key]:
                                if isinstance(existing_item, dict):
                                    merge_recursive(existing_item, list_item, source_guid)
                                    found = True
                                    break
                            if not found:
                                new_item = {}
                                merge_recursive(new_item, list_item, source_guid)
                                target[key].append(new_item)
                        else:
                            entry = {'guid': source_guid, 'value': list_item}
                            if entry not in target[key]:
                                target[key].append(entry)
                else:
                    if not isinstance(target[key], list):
                        target[key] = []
                    entry = {'guid': source_guid, 'value': value}
                    if entry not in target[key]:
                        target[key].append(entry)

    # 合并所有数据（跳过 DecLists）
    for data_with_guid in json_list_with_guid:
        json_data = data_with_guid.get('analysis', {}) or data_with_guid
        file_guid = data_with_guid.get('guid', str(uuid.uuid4()))
        merge_recursive(result, json_data, file_guid)

    # 第四步：构建新的 DecLists
    if dec_list_groups:
        if 'DecMessage' not in result:
            result['DecMessage'] = {}
        if 'DecLists' not in result['DecMessage']:
            result['DecMessage']['DecLists'] = {}

        # 为每个商品组合并数据，按照发票的 GNo 排序
        merged_dec_list = []
        sorted_groups = sorted(
            dec_list_groups.items(),
            key=lambda x: group_invoice_gno.get(x[0], float('inf'))
        )

        for goods_key, items_with_guid in sorted_groups:
            if items_with_guid:
                first_item = items_with_guid[0]['item']
                field_order = list(first_item.keys())
            else:
                field_order = []

            merged_item = {}
            for item_with_guid in items_with_guid:
                item_guid = item_with_guid['guid']
                item_data = item_with_guid['item']

                for key, value in item_data.items():
                    if key not in merged_item:
                        merged_item[key] = []

                    if isinstance(value, (dict, list)):
                        entry = {'guid': item_guid, 'value': value}
                        if entry not in merged_item[key]:
                            merged_item[key].append(entry)
                    else:
                        entry = {'guid': item_guid, 'value': value}
                        if entry not in merged_item[key]:
                            merged_item[key].append(entry)

            # 按原始字段顺序重新排列
            ordered_item = {}
            for key in field_order:
                if key in merged_item:
                    ordered_item[key] = merged_item[key]
            for key in merged_item:
                if key not in ordered_item:
                    ordered_item[key] = merged_item[key]

            merged_dec_list.append(ordered_item)

        result['DecMessage']['DecLists']['DecList'] = merged_dec_list

        # 第五步：应用企业默认值到所有商品
        if default_data:
            default_json = default_data['json']
            default_guid = default_data['guid']
            if 'DecMessage' in default_json and 'DecLists' in default_json['DecMessage']:
                default_dec_list = default_json['DecMessage']['DecLists'].get('DecList', [])
                if isinstance(default_dec_list, list) and len(default_dec_list) > 0:
                    default_item = default_dec_list[0]
                    log_message(f"应用企业默认值到所有商品: {list(default_item.keys())}")

                    for item in result['DecMessage']['DecLists']['DecList']:
                        for key, value in default_item.items():
                            if key == 'GNo':
                                continue
                            if key not in item or len(item[key]) == 0:
                                item[key] = []
                                entry = {'guid': default_guid, 'value': value}
                                item[key].append(entry)

    # 第六步：过滤掉不需要的值
    def filter_values(obj, parent_key=None):
        if isinstance(obj, dict):
            for key, value in obj.items():
                obj[key] = filter_values(value, key)
        elif isinstance(obj, list):
            if parent_key == 'FileBusinessType':
                return obj
            if len(obj) > 0 and isinstance(obj[0], dict) and 'guid' in obj[0] and 'value' in obj[0]:
                all_values = [item['value'] for item in obj]

                all_empty_or_zero = True
                for val in all_values:
                    if val is not None and val != '' and not (isinstance(val, (int, float)) and val == 0):
                        all_empty_or_zero = False
                        break

                if all_empty_or_zero:
                    return []

                filtered = []
                for item in obj:
                    if should_include_value(item['value'], all_values):
                        filtered.append(item)
                return filtered
            else:
                return [filter_values(item) for item in obj]
        return obj

    result = filter_values(result)
    result['FileBusinessType'] = '正式报关单'

    return result, guid_to_business_type


def find_best_match_by_description(new_desc, existing_items, threshold=70):
    """只根据 Description 匹配"""
    best_match = None
    best_score = 0

    for existing_item in existing_items:
        score = calculate_description_similarity(new_desc, existing_item['desc'])
        if score > best_score:
            best_score = score
            best_match = existing_item['key']

    if best_score >= threshold:
        return best_match, best_score

    return None, best_score


def calculate_description_similarity(s1, s2):
    """计算描述相似度"""
    if not s1 or not s2:
        return 0

    score1 = fuzz.token_sort_ratio(s1, s2)
    score2 = fuzz.partial_ratio(s1, s2)
    score3 = fuzz.ratio(s1, s2)

    return max(score1, score2, score3)


def should_include_value(value, all_values):
    """判断是否应该包含该值"""
    has_non_empty = False
    has_non_zero = False

    for val in all_values:
        if val is not None and val != '':
            has_non_empty = True
        if isinstance(val, (int, float)) and val != 0:
            has_non_zero = True

    if value is None or value == '':
        return not has_non_empty
    if isinstance(value, (int, float)) and value == 0:
        return not has_non_zero

    return True


def get_value_from_merged_data(merged_data, field_path):
    """从合并后的数据获取字段值"""
    parts = field_path.split('.')
    current = merged_data

    for part in parts:
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return []

    if isinstance(current, list) and len(current) > 0 and isinstance(current[0], dict) and 'guid' in current[0] and 'value' in current[0]:
        return current

    return [{'guid': None, 'value': current}]


def query_product_info(manufacturer=None, english_desc=None, script_dir=None):
    """从向量数据库查询产品信息"""
    if not LANCEDB_AVAILABLE:
        log_message("向量数据库不可用，跳过产品查询", "WARNING")
        return None

    if not script_dir:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    persist_directory = os.path.join(script_dir, 'lance_db')

    if not os.path.exists(persist_directory):
        log_message(f"向量数据库目录不存在: {persist_directory}", "WARNING")
        return None

    try:
        query_text = english_desc
        results = query_from_db(
            persist_directory=persist_directory,
            collection_name='product_library',
            query_text=query_text,
            n_results=10,
            strict_match_fields=['制造企业'] if manufacturer else None
        )

        if results and english_desc:
            results = reorder_results_with_keyword_match(
                results, english_desc, ['英文描述']
            )

        if results and results.get('ids') and len(results['ids'][0]) > 0:
            best_meta = None
            best_distance = float('inf')

            for i in range(len(results['ids'][0])):
                distance = results['distances'][0][i] if results.get('distances') and len(results['distances'][0]) > i else None
                meta = results['metadatas'][0][i] if results.get('metadatas') and len(results['metadatas'][0]) > i else None

                if distance is not None and distance < 1.2 and distance < best_distance:
                    best_distance = distance
                    best_meta = meta
                    log_message(f"候选匹配 - 距离: {distance:.4f}, 元数据: {meta}")

            if best_meta:
                log_message(f"找到最佳匹配产品: {best_meta}")
                return best_meta

        return None
    except Exception as e:
        log_message(f"产品查询失败: {e}", "ERROR")
        return None


def is_empty_value(val):
    """检查值是否为空（包括空列表、空字典、0等）"""
    if val is None or val == '' or val == [] or val == {}:
        return True
    if isinstance(val, (int, float)) and val == 0:
        return True
    return False


def normalize_string(val):
    """去除字符串的前后空白符，并将中间的连续空格替换成一个空格"""
    if isinstance(val, str):
        val = val.strip()
        val = ' '.join(val.split())
    return val


def get_values_from_field(field_data, guid_to_business_type=None, allowed_business_types=None):
    """
    从字段数据中提取所有值
    如果提供了 guid_to_business_type 和 allowed_business_types，只返回允许业务类型的值
    """
    if not isinstance(field_data, list):
        return []
    
    values = []
    
    for item in field_data:
        if not isinstance(item, dict) or 'value' not in item:
            continue
            
        val = item['value']
        
        # 检查值是否为空
        if is_empty_value(val):
            continue
        
        # 规范化字符串值
        val = normalize_string(val)
        
        # 如果提供了业务类型筛选，检查是否允许
        if guid_to_business_type and allowed_business_types:
            guid = item.get('guid')
            if guid:
                business_type = guid_to_business_type.get(guid)
                if business_type not in allowed_business_types:
                    continue
        
        values.append(val)
    
    return values


def get_value_from_field_by_business_type(field_data, guid_to_business_type, business_type):
    """从字段数据中获取指定业务类型的值"""
    if isinstance(field_data, list):
        for item in field_data:
            if isinstance(item, dict) and 'value' in item:
                guid = item.get('guid')
                if guid and guid_to_business_type.get(guid) == business_type:
                    val = item['value']
                    if val is not None and val != '' and not (isinstance(val, (int, float)) and val == 0):
                        # 去除字符串的前后空白符，并将中间的连续空格替换成一个空格
                        if isinstance(val, str):
                            val = val.strip()
                            # 将连续的空格替换成一个空格
                            val = ' '.join(val.split())
                        return val
    return None


def parse_template_placeholder(text):
    """
    解析模板占位符
    返回: (is_optional, placeholders)
        is_optional: 是否是可选的（用 [] 包裹）
        placeholders: 占位符列表，例如 ["DecMessage.DecHead.TradeName", "DecMessage.DecHead.OwnerName"]
    """
    if not text:
        return False, []
    
    # 检查是否是可选的（用 [] 包裹）
    is_optional = False
    content = text
    
    if text.startswith('[') and text.endswith(']'):
        is_optional = True
        content = text[1:-1]
    
    # 提取所有 {} 内的占位符（使用预编译的正则表达式）
    placeholders = PLACEHOLDER_PATTERN.findall(content)
    
    return is_optional, placeholders


def get_top_left_cell(sheet, row, col):
    """
    获取合并单元格的左上角单元格
    如果不是合并单元格，返回原单元格
    """
    # 检查是否是合并单元格
    for merged_range in sheet.merged_cells.ranges:
        if row >= merged_range.min_row and row <= merged_range.max_row and \
           col >= merged_range.min_col and col <= merged_range.max_col:
            # 返回左上角单元格
            return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
    
    # 不是合并单元格
    return sheet.cell(row=row, column=col)


def fill_cell_with_values(sheet, row, col, values, original_text=None, placeholder=None, guid_to_business_type=None, helper_sheet=None, helper_row=None):
    """
    填写单元格
    - 如果值唯一，直接填入
    - 如果值不唯一，创建下拉框，字体用红色
    - 如果提供了 original_text 和 placeholder，则进行文本替换
    - 如果提供了 helper_sheet，则使用隐藏辅助工作表存储下拉选项
    - 保留单元格的原始格式，只修改字体颜色
    """
    if not OPENPYXL_AVAILABLE:
        return helper_row
    
    # 获取实际要操作的单元格（处理合并单元格）
    cell = get_top_left_cell(sheet, row, col)
    
    # 去重 - 使用字典保持顺序的同时去重，性能更好
    seen = set()
    unique_values = []
    
    def make_hashable(v):
        """将值转换为可哈希类型"""
        if isinstance(v, list):
            return tuple(v)
        return v
    
    for v in values:
        hashable_v = make_hashable(v)
        if hashable_v not in seen:
            seen.add(hashable_v)
            unique_values.append(v)
    
    # 确定字体颜色
    # - 多来源且值一致：黑色
    # - 多来源且值不一致：红色
    # - 单一来源：蓝色
    font_color = "000000"  # 默认黑色
    
    # 计算来源数量（如果提供了 guid_to_business_type）
    source_count = len(values) if values else 0
    
    if source_count > 1:
        if len(unique_values) == 1:
            # 多来源但值一致，黑色
            font_color = "000000"
        else:
            # 多来源且值不一致，红色
            font_color = "FF0000"
    elif source_count == 1:
        # 单一来源，蓝色
        font_color = "0000FF"
    
    # 处理占位符替换的辅助函数
    def replace_placeholders_smart(text, placeholder_values):
        """
        智能替换占位符，按要求处理空内容
        placeholder_values: {placeholder: value} 字典，有值的占位符
        """
        if not text:
            return text
        
        is_wrapped_in_brackets = text.startswith('[') and text.endswith(']')
        inner_text = text[1:-1] if is_wrapped_in_brackets else text
        
        # 提取所有占位符（使用预编译的正则表达式）
        all_placeholders = PLACEHOLDER_PATTERN.findall(inner_text)
        
        # 检查是否有占位符有值
        has_any_value = any(
            ph in placeholder_values and placeholder_values[ph] is not None and placeholder_values[ph] != ""
            for ph in all_placeholders
        )
        
        # 2.1) 只有占位符 [] 包裹的内容，内容全部为空的时候才需要清空 [] 及包裹的所有内容
        if is_wrapped_in_brackets and not has_any_value:
            return ""
        
        # 否则进行替换
        result = inner_text
        
        # 替换有值的占位符
        for ph, val in placeholder_values.items():
            if val is not None and val != "":
                result = result.replace(f"{{{ph}}}", str(val))
        
        # 清空没有识别到的内容及对应的 {}
        result = PLACEHOLDER_PATTERN.sub('', result)
        
        # 清理残留的分隔符（使用预编译的正则表达式）
        result = MULTI_PIPE_PATTERN.sub('|', result)
        result = PIPE_START_PATTERN.sub('', result)
        result = PIPE_END_PATTERN.sub('', result)
        
        # 处理 / 分隔符 - 特殊处理：保留开头和结尾的单个 /
        # 先合并连续的 /
        result = MULTI_SLASH_PATTERN.sub('/', result)
        # 注意：不要去掉开头和结尾的 /，因为像 /300/总价 这样的格式是期望的
        
        # 2.1) 否则只清空 []
        if is_wrapped_in_brackets:
            result = result.strip()
            return result if result else ""
        
        return result
    
    # 保存原始字体的其他属性
    original_font = cell.font
    new_font = None
    if original_font:
        # 复制原始字体属性，只修改颜色
        new_font = Font(
            name=original_font.name,
            size=original_font.size,
            bold=original_font.bold,
            italic=original_font.italic,
            vertAlign=original_font.vertAlign,
            underline=original_font.underline,
            strike=original_font.strike,
            color=font_color
        )
    else:
        # 没有原始字体，只设置颜色
        new_font = Font(color=font_color)
    
    if len(unique_values) == 1:
        # 唯一值
        if original_text and placeholder:
            # 替换占位符并清理其他占位符
            placeholder_values = {placeholder: unique_values[0]}
            cell.value = replace_placeholders_smart(original_text, placeholder_values)
        else:
            # 直接填入
            cell.value = unique_values[0]
        
        # 设置字体颜色，保留其他格式
        cell.font = new_font
    else:
        # 多个值
        if original_text and placeholder:
            # 用第一个值替换占位符并清理其他占位符
            placeholder_values = {placeholder: unique_values[0] if unique_values else ""}
            cell.value = replace_placeholders_smart(original_text, placeholder_values)
        else:
            cell.value = unique_values[0] if unique_values else ""
        
        # 设置字体颜色，保留其他格式
        cell.font = new_font
        
        # 创建数据验证（下拉框）- 使用隐藏辅助工作表方式
        try:
            if helper_sheet is not None and helper_row is not None:
                # 将选项写入辅助工作表
                start_col = 1
                for i, v in enumerate(unique_values):
                    helper_sheet.cell(row=helper_row, column=start_col + i).value = str(v)
                
                # 创建单元格范围引用
                end_col = start_col + len(unique_values) - 1
                range_str = f"'{helper_sheet.title}'!${openpyxl.utils.get_column_letter(start_col)}${helper_row}:${openpyxl.utils.get_column_letter(end_col)}${helper_row}"
                
                # 创建数据验证
                dv = DataValidation(type="list", formula1=range_str, allow_blank=True)
                sheet.add_data_validation(dv)
                dv.add(cell)
                
                # 辅助工作表行号加1
                return helper_row + 1
        except Exception as e:
            log_message(f"创建下拉框失败: {e}", "WARNING")
    
    return helper_row


def clear_cell_placeholder(sheet, row, col, placeholder):
    """清除单元格中的占位符"""
    cell = get_top_left_cell(sheet, row, col)
    if cell.value and str(cell.value) == placeholder:
        cell.value = ""


def fill_field_to_cell(sheet, row, col, field_data, guid_to_business_type, helper_sheet, helper_row, use_sum=False):
    """
    通用函数：将字段数据填写到单元格
    
    参数:
        sheet: 工作表
        row: 行号
        col: 列号
        field_data: 字段数据
        guid_to_business_type: GUID到业务类型的映射
        helper_sheet: 辅助工作表
        helper_row: 辅助行号
        use_sum: 是否使用求和（用于Qty字段）
    
    返回:
        新的helper_row或None
    """
    if use_sum and isinstance(field_data, list):
        values = get_qty_sum_by_guid(field_data, guid_to_business_type)
    else:
        values = get_values_from_field(field_data, guid_to_business_type) if isinstance(field_data, list) else []
    
    if values:
        return fill_cell_with_values(sheet, row, col, values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
    return None


def get_qty_sum_by_guid(qty_values, guid_to_business_type):
    """
    根据单据业务类型对应的 GUID 累加数量
    返回: 累加后的值列表
    """
    # 按业务类型分组累加
    business_type_sums = defaultdict(float)
    
    for item in qty_values:
        guid = item.get('guid')
        value = item.get('value')
        
        if guid and guid in guid_to_business_type:
            business_type = guid_to_business_type[guid]
            if value is not None:
                try:
                    business_type_sums[business_type] += float(value)
                except (ValueError, TypeError):
                    pass
    
    # 转换为列表
    sums = list(business_type_sums.values())
    
    # 如果没有累加结果，返回原始值
    if not sums:
        return [item.get('value') for item in qty_values]
    
    return sums


def fill_item_to_rows(sheet, item, start_row, guid_to_business_type, manufacturer=None, helper_sheet=None, helper_row=None):
    """
    将一个商品的数据填写到从 start_row 开始的 3 行
    """
    log_message(f"fill_item_to_rows 开始 - 起始行: {start_row}, manufacturer: {manufacturer}")
    log_message(f"商品数据 item: {item}")
    
    # 从商品中获取 Description（用于向量检索）- 只从发票获取
    english_desc = get_value_from_field_by_business_type(item.get('Description', []), guid_to_business_type, '发票')
    log_message(f"获取到的英文描述: {english_desc}")
    
    # 查询向量数据库
    vector_match = None
    if manufacturer and english_desc:
        log_message(f"查询向量数据库 - 制造企业: {manufacturer}, 英文描述: {english_desc}")
        vector_match = query_product_info(manufacturer, english_desc)
        log_message(f"向量查询结果 vector_match: {vector_match}")
    else:
        log_message(f"跳过向量查询 - manufacturer: {manufacturer}, english_desc: {english_desc}")
    
    # 第 1 行
    row1 = start_row
    
    # GNo（第2列）- 以发票为准
    gno_value = get_value_from_field_by_business_type(item.get('GNo', []), guid_to_business_type, '发票')
    if gno_value:
        new_helper_row = fill_cell_with_values(sheet, row1, 2, [gno_value], guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        clear_cell_placeholder(sheet, row1, 2, '{GNo}')
    
    # 税号（第3列）- 使用向量数据库
    cell = get_top_left_cell(sheet, row1, 3)
    if cell.value and str(cell.value) == '<税号>':
        cell.value = vector_match.get('税号', '') if vector_match else ''
        if vector_match and vector_match.get('税号'):
            log_message(f"填写税号: {vector_match['税号']}")
    
    # 品名（第5列前半部分）- 使用向量数据库
    cell = get_top_left_cell(sheet, row1, 5)
    if cell.value and str(cell.value) == '<品名>':
        cell.value = vector_match.get('品名', '') if vector_match else ''
        if vector_match and vector_match.get('品名'):
            log_message(f"填写品名: {vector_match['品名']}")
    
    # Description（第11列）- 从所有数据源获取
    new_helper_row = fill_field_to_cell(sheet, row1, 11, item.get('Description', []), guid_to_business_type, helper_sheet, helper_row)
    if new_helper_row is not None:
        helper_row = new_helper_row
    else:
        clear_cell_placeholder(sheet, row1, 11, '{Description}')
    
    # Qty1（第16列）
    new_helper_row = fill_field_to_cell(sheet, row1, 16, item.get('Qty1', []), guid_to_business_type, helper_sheet, helper_row, use_sum=True)
    if new_helper_row is not None:
        helper_row = new_helper_row
    else:
        clear_cell_placeholder(sheet, row1, 16, '{Qty1}')
    
    # Unit1（第19列）
    new_helper_row = fill_field_to_cell(sheet, row1, 19, item.get('Unit1', []), guid_to_business_type, helper_sheet, helper_row)
    if new_helper_row is not None:
        helper_row = new_helper_row
    else:
        clear_cell_placeholder(sheet, row1, 19, '{Unit1}')
    
    # DeclTotal（第20列）
    new_helper_row = fill_field_to_cell(sheet, row1, 20, item.get('DeclTotal', []), guid_to_business_type, helper_sheet, helper_row)
    if new_helper_row is not None:
        helper_row = new_helper_row
    else:
        clear_cell_placeholder(sheet, row1, 20, '{DeclTotal}')
    
    # DestinationCountry（第25列）
    dest_country_values = get_values_from_field(item.get('DestinationCountry', []))
    if dest_country_values:
        new_helper_row = fill_cell_with_values(sheet, row1, 25, dest_country_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row1, 25)
        if cell.value and str(cell.value) == '{DestinationCountry}':
            cell.value = ''
    
    # OriginCountry（第23列）
    origin_country_values = get_values_from_field(item.get('OriginCountry', []))
    if origin_country_values:
        new_helper_row = fill_cell_with_values(sheet, row1, 23, origin_country_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row1, 23)
        if cell.value and str(cell.value) == '{OriginCountry}':
            cell.value = ''
    
    # District（第30列）
    district_values = get_values_from_field(item.get('District', []))
    if district_values:
        new_helper_row = fill_cell_with_values(sheet, row1, 30, district_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row1, 30)
        if cell.value and str(cell.value) == '{District}':
            cell.value = ''
    
    # DutyMode（第35列）
    duty_mode_values = get_values_from_field(item.get('DutyMode', []))
    if duty_mode_values:
        new_helper_row = fill_cell_with_values(sheet, row1, 35, duty_mode_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row1, 35)
        if cell.value and str(cell.value) == '{DutyMode}':
            cell.value = ''
    
    # 第 2 行
    row2 = start_row + 1
    
    # ContrItem（第2列）
    contr_item_values = get_values_from_field(item.get('ContrItem', []))
    if contr_item_values:
        new_helper_row = fill_cell_with_values(sheet, row2, 2, [f"({v})" for v in contr_item_values], guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row2, 2)
        if cell.value and str(cell.value) == '(ContrItem)':
            cell.value = ''
    
    # 申报要素（第5列）- 使用向量数据库
    cell = get_top_left_cell(sheet, row2, 5)
    if cell.value and str(cell.value) == '<申报要素>':
        if vector_match and vector_match.get('申报要素'):
            cell.value = vector_match['申报要素']
            log_message(f"填写申报要素: {vector_match['申报要素']}")
        else:
            cell.value = ''
    
    # Qty2 / NetWt（第16列）
    qty2_values = item.get('Qty2', [])
    if qty2_values:
        sums = get_qty_sum_by_guid(qty2_values, guid_to_business_type)
        new_helper_row = fill_cell_with_values(sheet, row2, 16, sums, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        # 尝试填写 NetWt
        netwt_values = get_values_from_field(item.get('NetWt', []))
        if netwt_values:
            new_helper_row = fill_cell_with_values(sheet, row2, 16, netwt_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
            if new_helper_row is not None:
                helper_row = new_helper_row
        else:
            cell = get_top_left_cell(sheet, row2, 16)
            if cell.value and (str(cell.value) == '{Qty2}' or str(cell.value) == '{NetWt}'):
                cell.value = ''
    
    # Unit2（第19列）
    unit2_values = get_values_from_field(item.get('Unit2', []))
    if unit2_values:
        new_helper_row = fill_cell_with_values(sheet, row2, 19, unit2_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row2, 19)
        if cell.value and str(cell.value) == '{Unit2}':
            cell.value = ''
    
    # TradeCurr（第20列）
    trade_curr_values = get_values_from_field(item.get('TradeCurr', []))
    if trade_curr_values:
        new_helper_row = fill_cell_with_values(sheet, row2, 20, trade_curr_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row2, 20)
        if cell.value and str(cell.value) == '{TradeCurr}':
            cell.value = ''
    
    # 第 3 行
    row3 = start_row + 2
    
    # Qty3（第16列）
    qty3_values = item.get('Qty3', [])
    if qty3_values:
        sums = get_qty_sum_by_guid(qty3_values, guid_to_business_type)
        new_helper_row = fill_cell_with_values(sheet, row3, 16, sums, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row3, 16)
        if cell.value and str(cell.value) == '{Qty3}':
            cell.value = ''
    
    # Unit3（第19列）
    unit3_values = get_values_from_field(item.get('Unit3', []))
    if unit3_values:
        new_helper_row = fill_cell_with_values(sheet, row3, 19, unit3_values, guid_to_business_type=guid_to_business_type, helper_sheet=helper_sheet, helper_row=helper_row)
        if new_helper_row is not None:
            helper_row = new_helper_row
    else:
        cell = get_top_left_cell(sheet, row3, 19)
        if cell.value and str(cell.value) == '{Unit3}':
            cell.value = ''
    
    return helper_row


def fill_declaration_excel(template_path, merged_data, output_dir, guid_to_business_type=None):
    """
    填写报关单 Excel
    """
    if not OPENPYXL_AVAILABLE:
        log_message("openpyxl 未安装，无法生成 Excel", "ERROR")
        return None

    if not os.path.exists(template_path):
        log_message(f"模板文件不存在: {template_path}", "ERROR")
        return None

    os.makedirs(output_dir, exist_ok=True)

    try:
        log_message(f"开始填写报关单: {template_path}")
        
        # 加载工作簿
        wb = openpyxl.load_workbook(template_path)
        
        # 创建隐藏的辅助工作表，用于存储下拉选项
        helper_sheet = wb.create_sheet(title="_dropdown_options")
        helper_sheet.sheet_state = "hidden"  # 隐藏工作表
        helper_row = 1  # 从第一行开始
        
        # 处理每个工作表
        for sheet in wb.worksheets:
            if sheet.title == "_dropdown_options":
                continue  # 跳过辅助工作表
            
            log_message(f"处理工作表: {sheet.title}")
            
            # 第一步：获取发票的境内收发货人名称作为制造企业
            manufacturer = None
            trade_name_values = get_value_from_merged_data(merged_data, "DecMessage.DecHead.TradeName")
            if trade_name_values:
                for item in trade_name_values:
                    value = item.get('value')
                    guid = item.get('guid')
                    # 只从发票获取
                    if guid and guid_to_business_type.get(guid) == '发票' and value is not None and value != '':
                        manufacturer = value
                        log_message(f"使用制造企业（来自发票）: {manufacturer}")
                        break
            
            # 第二步：只处理表头（行 1-22）的占位符，不处理表体
            # 记录已处理的合并单元格左上角坐标，避免重复处理
            log_message("开始处理表头...")
            processed_cells = set()
            
            for row_idx in range(1, 23):
                for col_idx in range(1, sheet.max_column + 1):
                    # 直接获取左上角单元格，避免重复处理
                    top_left = get_top_left_cell(sheet, row_idx, col_idx)
                    cell_key = (top_left.row, top_left.column)
                    
                    if cell_key in processed_cells:
                        continue
                    
                    if not top_left.value or not isinstance(top_left.value, str):
                        continue
                    
                    cell_value = str(top_left.value)
                    log_message(f"检查单元格 ({row_idx},{col_idx}) - 原始值: '{cell_value}'")
                    is_optional, placeholders = parse_template_placeholder(cell_value)
                    log_message(f"  解析结果: is_optional={is_optional}, placeholders={placeholders}")
                    
                    if not placeholders:
                        continue
                    
                    processed_cells.add(cell_key)
                    
                    # 收集所有占位符的值
                    placeholder_values = {}  # {placeholder: (first_value, all_values)}
                    has_any_value = False
                    
                    for placeholder in placeholders:
                        full_field_path = FIELD_PATH_MAP.get(placeholder, placeholder)
                        values = get_value_from_merged_data(merged_data, full_field_path)
                        
                        # 使用 get_values_from_field 提取所有非空值
                        field_values = get_values_from_field(values)
                        
                        if field_values:
                            # 保存第一个值用于显示，以及所有值用于可能的下拉框
                            placeholder_values[placeholder] = (field_values[0], field_values)
                            has_any_value = True
                        else:
                            placeholder_values[placeholder] = (None, [])
                    
                    # 构建只包含第一个值的字典，用于占位符替换
                    placeholder_first_values = {ph: val[0] for ph, val in placeholder_values.items()}
                    
                    # 先使用智能替换设置单元格值（无论是否有值都需要处理）
                    is_wrapped_in_brackets = cell_value.startswith('[') and cell_value.endswith(']')
                    
                    # 使用新的智能替换函数
                    def smart_replace():
                        if not cell_value:
                            return ""
                        
                        inner_text = cell_value[1:-1] if is_wrapped_in_brackets else cell_value
                        
                        # 检查是否所有占位符都为空
                        all_empty = not any(
                            placeholder_first_values.get(ph) is not None and placeholder_first_values.get(ph) != ""
                            for ph in placeholders
                        )
                        
                        # 2.1) 只有占位符 [] 包裹的内容，内容全部为空的时候才需要清空 [] 及包裹的所有内容
                        if is_wrapped_in_brackets and all_empty:
                            return ""
                        
                        # 否则进行替换
                        result = inner_text
                        
                        # 替换有值的占位符
                        for ph, val in placeholder_first_values.items():
                            if val is not None and val != "":
                                result = result.replace(f"{{{ph}}}", str(val))
                        
                        # 2.2) 清空没有识别到的内容及对应的 {}（使用预编译的正则表达式）
                        result = PLACEHOLDER_PATTERN.sub('', result)
                        
                        # 清理残留的分隔符（使用预编译的正则表达式）
                        result = MULTI_PIPE_PATTERN.sub('|', result)
                        result = PIPE_START_PATTERN.sub('', result)
                        result = PIPE_END_PATTERN.sub('', result)
                        
                        # 处理 / 分隔符 - 特殊处理：保留开头和结尾的单个 /
                        # 先合并连续的 /
                        result = MULTI_SLASH_PATTERN.sub('/', result)
                        # 注意：不要去掉开头和结尾的 /，因为像 /300/总价 这样的格式是期望的
                        
                        # 2.1) 否则只清空 []
                        if is_wrapped_in_brackets:
                            result = result.strip()
                            return result if result else ""
                        
                        return result
                    
                    # 设置单元格值（无论是否有值都需要设置）
                    top_left.value = smart_replace()
                    
                    # 只有当有值时才设置字体颜色和下拉框
                    if has_any_value:
                        # 找到第一个有值的占位符，用于确定颜色和下拉框
                        first_ph_with_value = None
                        first_field_values = []
                        for placeholder in placeholders:
                            if placeholder_values[placeholder][0] is not None:
                                first_ph_with_value = placeholder
                                first_field_values = placeholder_values[placeholder][1]
                                break
                        
                        if first_ph_with_value:
                            # 保存原始字体的其他属性
                            original_font = top_left.font
                            
                            # 确定字体颜色
                            source_count = len(first_field_values) if first_field_values else 0
                            unique_values = []
                            seen = set()
                            for v in first_field_values:
                                if isinstance(v, list):
                                    v_tuple = tuple(v)
                                    if v_tuple not in seen:
                                        seen.add(v_tuple)
                                        unique_values.append(v)
                                else:
                                    if v not in seen:
                                        seen.add(v)
                                        unique_values.append(v)
                            
                            font_color = "000000"
                            if source_count > 1:
                                if len(unique_values) == 1:
                                    font_color = "000000"
                                else:
                                    font_color = "FF0000"
                            elif source_count == 1:
                                font_color = "0000FF"
                            
                            # 复制原始字体属性，只修改颜色
                            new_font = None
                            if original_font:
                                new_font = Font(
                                    name=original_font.name,
                                    size=original_font.size,
                                    bold=original_font.bold,
                                    italic=original_font.italic,
                                    vertAlign=original_font.vertAlign,
                                    underline=original_font.underline,
                                    strike=original_font.strike,
                                    color=font_color
                                )
                            else:
                                new_font = Font(color=font_color)
                            
                            top_left.font = new_font
                            
                            # 如果有多个值，创建下拉框
                            if len(unique_values) > 1 and helper_sheet is not None and helper_row is not None:
                                try:
                                    # 将选项写入辅助工作表
                                    start_col = 1
                                    for i, v in enumerate(unique_values):
                                        helper_sheet.cell(row=helper_row, column=start_col + i).value = str(v)
                                    
                                    # 创建单元格范围引用
                                    end_col = start_col + len(unique_values) - 1
                                    range_str = f"'{helper_sheet.title}'!${openpyxl.utils.get_column_letter(start_col)}${helper_row}:${openpyxl.utils.get_column_letter(end_col)}${helper_row}"
                                    
                                    # 创建数据验证
                                    dv = DataValidation(type="list", formula1=range_str, allow_blank=True)
                                    sheet.add_data_validation(dv)
                                    dv.add(top_left)
                                    
                                    helper_row += 1
                                except Exception as e:
                                    log_message(f"创建下拉框失败: {e}", "WARNING")
                    
                    log_message(f"  最终处理完成")
            
            # 第三步：获取表体数据
            dec_list_data = []
            if 'DecMessage' in merged_data and \
               'DecLists' in merged_data['DecMessage'] and \
               'DecList' in merged_data['DecMessage']['DecLists']:
                dec_list_data = merged_data['DecMessage']['DecLists']['DecList']
            
            log_message(f"表体数据行数: {len(dec_list_data)}")
            
            # 第四步：先填写商品数据（从第23行开始，每个对象3行）
            if dec_list_data:
                log_message(f"开始处理 {len(dec_list_data)} 个商品")
                
                # 逐个填写商品数据
                for item_idx, item in enumerate(dec_list_data):
                    base_row = 23 + 3 * item_idx
                    log_message(f"填写商品 {item_idx + 1}，起始行: {base_row}")
                    helper_row = fill_item_to_rows(
                        sheet, item, base_row, guid_to_business_type, manufacturer,
                        helper_sheet=helper_sheet, helper_row=helper_row
                    )
            
            # 第五步：找到特殊关系确认行
            special_relation_row = None
            max_col = sheet.max_column
            for row_idx in range(1, sheet.max_row + 1):
                for col_idx in range(1, max_col + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    if cell.value and "特殊关系确认" in str(cell.value):
                        special_relation_row = row_idx
                        log_message(f"找到特殊关系确认行: {special_relation_row}")
                        break
                if special_relation_row:
                    break
            
            # 第六步：删除特殊关系行之上，行对象填到的表格行之下的所有行
            if dec_list_data and special_relation_row:
                # 计算需要删除的起始行：从 (23 + 3 * len(dec_list_data)) 开始，到特殊关系行 - 1 结束
                start_delete_row = 23 + 3 * len(dec_list_data)
                
                if start_delete_row < special_relation_row:
                    rows_to_delete = special_relation_row - start_delete_row
                    log_message(f"删除 {rows_to_delete} 行多余内容（从第 {start_delete_row} 行到第 {special_relation_row - 1} 行")
                    
                    # 先处理合并单元格：找出所有与要删除的行重叠的合并单元格
                    merged_ranges_to_remove = []
                    for merged_range in sheet.merged_cells.ranges:
                        # 检查合并单元格是否与要删除的行有重叠
                        if (merged_range.min_row >= start_delete_row and merged_range.max_row < special_relation_row) or \
                           (merged_range.max_row >= start_delete_row and merged_range.min_row < special_relation_row):
                            merged_ranges_to_remove.append(merged_range)
                    
                    # 取消这些合并单元格
                    for merged_range in merged_ranges_to_remove:
                        sheet.unmerge_cells(str(merged_range))
                        log_message(f"取消合并单元格: {merged_range}")
                    
                    # 从后往前删除，避免行号变化
                    for i in range(rows_to_delete):
                        sheet.delete_rows(special_relation_row - 1 - i)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 先保存合并结果为 JSON
        json_output_path = os.path.join(output_dir, f"declaration_merged_{timestamp}.json")
        with open(json_output_path, 'w', encoding='utf-8') as f:
            json.dump(merged_data, f, ensure_ascii=False, indent=2)
        log_message(f"合并数据已保存: {json_output_path}")
        
        # 保存 Excel
        excel_output_path = os.path.join(output_dir, f"报关单_{timestamp}.xlsx")
        wb.save(excel_output_path)
        log_message(f"报关单已生成: {excel_output_path}")
        
        return excel_output_path
    except Exception as e:
        log_message(f"生成 Excel 失败: {e}", "ERROR")
        import traceback
        log_message(f"异常堆栈:\n{traceback.format_exc()}", "DEBUG")
        return None


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="报关单生成器")
    parser.add_argument(
        "--analysis-results",
        help="分析结果 JSON 文件路径"
    )
    parser.add_argument(
        "--analysis-json",
        help="分析结果 JSON 字符串"
    )
    parser.add_argument(
        "--output-dir",
        default="/mnt/user-data/outputs",
        help="输出目录（默认：/mnt/user-data/outputs）"
    )
    parser.add_argument(
        "--template-path",
        help="报关单模板路径"
    )

    args = parser.parse_args()

    # 设置默认模板路径
    if not args.template_path:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        args.template_path = os.path.join(script_dir, 'assets', '报关单.xlsx')

    log_message("=" * 60)
    log_message("报关单生成器")
    log_message("=" * 60)

    # 加载分析结果
    analysis_results = load_analysis_results(args.analysis_results, args.analysis_json)
    if not analysis_results:
        log_message("无法加载分析结果，退出", "ERROR")
        return 1

    log_message(f"已加载 {len(analysis_results) if isinstance(analysis_results, list) else 1} 个分析结果")

    # 确保是列表格式
    if not isinstance(analysis_results, list):
        analysis_results = [analysis_results]

    # 为每个结果添加 GUID
    for i, result in enumerate(analysis_results):
        if 'guid' not in result:
            result['guid'] = str(uuid.uuid4())

    # 合并数据
    log_message("正在合并数据...")
    merged_data, guid_to_business_type = merge_json_objects_with_guid(analysis_results)

    # 尝试查询产品信息
    if 'DecMessage' in merged_data and 'DecLists' in merged_data['DecMessage']:
        dec_list = merged_data['DecMessage']['DecLists'].get('DecList', [])
        for item in dec_list:
            # 提取制造商和英文描述
            manufacturer = None
            english_desc = None

            # 尝试从合并数据中获取
            if 'DecHead' in merged_data.get('DecMessage', {}):
                trade_name = get_value_from_merged_data(merged_data, 'DecMessage.DecHead.TradeName')
                if trade_name and len(trade_name) > 0:
                    # 只从发票获取制造商
                    for item_guid in trade_name:
                        guid = item_guid.get('guid')
                        if guid and guid_to_business_type.get(guid) == '发票':
                            manufacturer = item_guid.get('value')
                            break

            desc_list = item.get('Description', [])
            if desc_list and len(desc_list) > 0:
                # 只从发票获取英文描述
                for item_guid in desc_list:
                    guid = item_guid.get('guid')
                    if guid and guid_to_business_type.get(guid) == '发票':
                        english_desc = item_guid.get('value')
                        break

            if english_desc:
                product_info = query_product_info(manufacturer, english_desc)
                if product_info:
                    log_message(f"为商品 '{english_desc}' 找到产品信息")
                    # 可以在这里将产品信息合并到 item 中

    # 生成 Excel
    log_message("正在生成报关单...")
    excel_path = fill_declaration_excel(
        args.template_path,
        merged_data,
        args.output_dir,
        guid_to_business_type
    )

    if excel_path:
        log_message("=" * 60)
        log_message("处理完成!")
        log_message(f"报关单: {excel_path}")
        log_message("=" * 60)
        return 0
    else:
        log_message("生成报关单失败", "ERROR")
        return 1


if __name__ == "__main__":
    sys.exit(main())
