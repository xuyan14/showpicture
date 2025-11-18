#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将Excel文件转换为JSON格式，用于前端预览
"""
import pandas as pd
import json
import os
import re
from pathlib import Path

def extract_number_from_title(title):
    """从标题中提取数字，例如 'VIP自A加绒4雪+6+腾讯...' 返回 6"""
    if not title or not isinstance(title, str):
        return 999  # 如果无法提取，放到最后
    # 匹配 +数字+ 的模式
    match = re.search(r'\+(\d+)\+', title)
    if match:
        return int(match.group(1))
    return 999  # 如果无法提取，放到最后

def sort_by_title_number(material_titles, material_urls, material_ids):
    """按照标题中的数字对三个数组进行排序"""
    if not material_titles or len(material_titles) == 0:
        return material_titles, material_urls, material_ids
    
    # 如果URL数量多于标题数量，需要为多余的URL创建虚拟标题
    # 确保所有URL都能被保留
    max_len = max(len(material_titles), len(material_urls), len(material_ids))
    
    # 扩展数组到相同长度
    extended_titles = list(material_titles) + [None] * (max_len - len(material_titles))
    extended_urls = list(material_urls) + [None] * (max_len - len(material_urls))
    extended_ids = list(material_ids) + [None] * (max_len - len(material_ids))
    
    # 创建索引和数字的配对
    indexed_items = []
    for i in range(max_len):
        title = extended_titles[i]
        url = extended_urls[i]
        id_val = extended_ids[i]
        
        # 如果有标题，提取数字；如果没有标题但有URL，使用999（放到最后）
        if title:
            num = extract_number_from_title(title)
        elif url:
            num = 999  # 没有标题的URL放到最后
        else:
            continue  # 跳过空项
        
        indexed_items.append({
            'index': i,
            'number': num,
            'title': title,
            'url': url,
            'id': id_val
        })
    
    # 按照数字排序
    indexed_items.sort(key=lambda x: x['number'])
    
    # 重新构建排序后的数组，过滤掉None值
    sorted_titles = [item['title'] for item in indexed_items if item['title'] is not None]
    sorted_urls = [item['url'] for item in indexed_items if item['url'] is not None]
    sorted_ids = [item['id'] for item in indexed_items if item['id'] is not None]
    
    return sorted_titles, sorted_urls, sorted_ids

def convert_excel_to_json(excel_path, output_path, data_source_name, category_id=None, layout=None):
    """将Excel或CSV文件转换为JSON格式
    
    Args:
        excel_path: Excel或CSV文件路径
        output_path: 输出JSON文件路径
        data_source_name: 数据源名称
        category_id: 类别ID，如果为None则根据数据源名称自动判断
        layout: 布局类型（'grid'或'strip'），如果为None则根据数据源名称自动判断
    """
    print(f"正在读取 {excel_path}...")
    # 支持CSV和Excel格式
    if excel_path.suffix.lower() == '.csv':
        df = pd.read_csv(excel_path)
    else:
        df = pd.read_excel(excel_path)
    print(f"读取到 {len(df)} 条记录")
    
    # 根据数据源名称自动判断类别ID和布局
    if category_id is None:
        if '多品多图' in data_source_name:
            category_id = 'multi'
            layout = layout or 'grid'
        elif '抖音图文' in data_source_name:
            category_id = 'douyin'
            layout = layout or 'strip'
        else:
            category_id = 'multi'
            layout = layout or 'grid'
    else:
        layout = layout or 'grid'
    
    items = []
    for idx, row in df.iterrows():
        # 解析JSON字符串字段
        material_ids = []
        material_titles = []
        material_urls = []
        
        # 解析material_id_list
        if pd.notna(row.get('material_id_list')):
            try:
                material_ids = json.loads(str(row['material_id_list']))
            except:
                material_ids = []
        
        # 处理标题列表（支持多种列名）
        # CSV格式: otd_material_title_list
        # Excel格式: corrected_title_list, original_title_list
        if pd.notna(row.get('corrected_title_list')):
            try:
                material_titles = json.loads(str(row['corrected_title_list']))
            except:
                material_titles = []
        
        if not material_titles and pd.notna(row.get('original_title_list')):
            try:
                material_titles = json.loads(str(row['original_title_list']))
            except:
                material_titles = []
        
        if not material_titles and pd.notna(row.get('otd_material_title_list')):
            try:
                title_str = str(row['otd_material_title_list'])
                # 修复CSV中的中文逗号问题
                title_str = title_str.replace('，', ',')  # 中文逗号替换为英文逗号
                material_titles = json.loads(title_str)
            except:
                # 如果JSON解析失败，尝试使用eval（仅用于CSV格式）
                try:
                    title_str = str(row['otd_material_title_list']).replace('，', ',')
                    material_titles = eval(title_str) if isinstance(eval(title_str), list) else []
                except:
                    material_titles = []
        
        # 处理URL列表（支持多种列名）
        # CSV格式: material_url_list
        # Excel格式: corrected_url_list, original_url_list
        if pd.notna(row.get('corrected_url_list')):
            try:
                url_str = str(row['corrected_url_list'])
                # 修复CSV中的中文逗号问题
                url_str = url_str.replace('，', ',')  # 中文逗号替换为英文逗号
                material_urls = json.loads(url_str)
            except:
                # 如果JSON解析失败，尝试使用eval（仅用于CSV格式）
                try:
                    url_str = str(row['corrected_url_list']).replace('，', ',')
                    parsed = eval(url_str)
                    material_urls = parsed if isinstance(parsed, list) else []
                except:
                    material_urls = []
            
            # 如果material_urls中的元素是字符串且包含逗号，需要分割
            if material_urls and isinstance(material_urls, list):
                expanded_urls = []
                for url_item in material_urls:
                    if isinstance(url_item, str):
                        # 如果字符串包含逗号，按逗号分割
                        if ',' in url_item:
                            split_urls = [u.strip() for u in url_item.split(',') if u.strip()]
                            expanded_urls.extend(split_urls)
                        else:
                            expanded_urls.append(url_item)
                    else:
                        expanded_urls.append(url_item)
                material_urls = expanded_urls
        
        if not material_urls and pd.notna(row.get('original_url_list')):
            try:
                url_str = str(row['original_url_list'])
                # 修复CSV中的中文逗号问题
                url_str = url_str.replace('，', ',')  # 中文逗号替换为英文逗号
                material_urls = json.loads(url_str)
            except:
                # 如果JSON解析失败，尝试使用eval（仅用于CSV格式）
                try:
                    url_str = str(row['original_url_list']).replace('，', ',')
                    parsed = eval(url_str)
                    material_urls = parsed if isinstance(parsed, list) else []
                except:
                    material_urls = []
            
            # 如果material_urls中的元素是字符串且包含逗号，需要分割
            if material_urls and isinstance(material_urls, list):
                expanded_urls = []
                for url_item in material_urls:
                    if isinstance(url_item, str):
                        # 如果字符串包含逗号，按逗号分割
                        if ',' in url_item:
                            split_urls = [u.strip() for u in url_item.split(',') if u.strip()]
                            expanded_urls.extend(split_urls)
                        else:
                            expanded_urls.append(url_item)
                    else:
                        expanded_urls.append(url_item)
                material_urls = expanded_urls
        
        if not material_urls and pd.notna(row.get('material_url_list')):
            try:
                url_str = str(row['material_url_list'])
                # 修复CSV中的中文逗号问题
                url_str = url_str.replace('，', ',')  # 中文逗号替换为英文逗号
                material_urls = json.loads(url_str)
            except:
                # 如果JSON解析失败，尝试使用eval（仅用于CSV格式）
                try:
                    url_str = str(row['material_url_list']).replace('，', ',')
                    parsed = eval(url_str)
                    material_urls = parsed if isinstance(parsed, list) else []
                except:
                    material_urls = []
            
            # 如果material_urls中的元素是字符串且包含逗号，需要分割
            if material_urls and isinstance(material_urls, list):
                expanded_urls = []
                for url_item in material_urls:
                    if isinstance(url_item, str):
                        # 如果字符串包含逗号，按逗号分割
                        if ',' in url_item:
                            split_urls = [u.strip() for u in url_item.split(',') if u.strip()]
                            expanded_urls.extend(split_urls)
                        else:
                            expanded_urls.append(url_item)
                    else:
                        expanded_urls.append(url_item)
                material_urls = expanded_urls
        
        # 按照标题中的数字（+1、+2、+3等）对所有数组进行排序
        # 只有当标题数量和URL数量匹配时，才进行排序
        # 如果标题数量少于URL数量，保持原始顺序（避免打乱URL顺序）
        if material_titles and len(material_titles) > 0:
            if len(material_titles) == len(material_urls):
                # 标题和URL数量匹配，可以安全排序
                material_titles, material_urls, material_ids = sort_by_title_number(
                    material_titles, material_urls, material_ids
                )
            elif len(material_titles) < len(material_urls):
                # 标题数量少于URL数量，只对标题部分排序，保持URL原始顺序
                # 这种情况通常出现在抖音图文数据中，标题可能只代表其中一个URL
                # 我们保持URL的原始顺序，只对标题进行排序
                sorted_titles, _, sorted_ids = sort_by_title_number(
                    material_titles, 
                    material_titles[:len(material_urls)] if len(material_titles) <= len(material_urls) else material_titles,
                    material_ids[:len(material_urls)] if len(material_ids) <= len(material_urls) else material_ids
                )
                material_titles = sorted_titles
                # URLs保持原始顺序，不排序
        
        # 构建记录
        record = {
            "report_date": str(int(row['report_date'])) if pd.notna(row.get('report_date')) else "",
            "unique_material_id": str(row['unique_material_id']) if pd.notna(row.get('unique_material_id')) else "(NULL)",
            "material_urls": material_urls if isinstance(material_urls, list) else [],
            "material_titles": material_titles if isinstance(material_titles, list) else [],
            "material_ids": material_ids if isinstance(material_ids, list) else [],
        }
        
        # 添加cost字段（如果存在）
        if pd.notna(row.get('cost')):
            record["cost"] = float(row['cost'])
        else:
            record["cost"] = 0.0
        
        items.append(record)
    
    # 去重：基于URL组合去重，保留第一条记录
    print(f"去重前记录数: {len(items)}")
    seen_urls = {}
    unique_items = []
    duplicates_count = 0
    
    for item in items:
        # 使用排序后的URL列表作为唯一标识
        urls = item.get('material_urls', [])
        url_key = tuple(sorted(urls)) if urls else tuple()
        
        if url_key and url_key in seen_urls:
            duplicates_count += 1
            # 如果当前记录的cost更高，则替换已存在的记录（可选策略）
            # 这里我们采用保留第一条的策略，如果需要保留cost最高的，可以取消下面的注释
            # existing_item = seen_urls[url_key]
            # if item.get('cost', 0) > existing_item.get('cost', 0):
            #     # 替换为cost更高的记录
            #     unique_items.remove(existing_item)
            #     unique_items.append(item)
            #     seen_urls[url_key] = item
        else:
            if url_key:  # 只处理有URL的记录
                seen_urls[url_key] = item
            unique_items.append(item)
    
    print(f"去重后记录数: {len(unique_items)}，去除了 {duplicates_count} 条重复记录")
    
    # 过滤掉无URL的记录
    before_filter = len(unique_items)
    unique_items = [
        item for item in unique_items
        if item.get('material_urls') and len(item.get('material_urls', [])) > 0
    ]
    filtered_count = before_filter - len(unique_items)
    if filtered_count > 0:
        print(f"过滤掉无URL的记录: {filtered_count} 条")
    
    # 构建输出格式
    output_data = {
        "categories": [
            {
                "id": category_id,
                "name": data_source_name,
                "layout": layout,
                "items": unique_items
            }
        ]
    }
    
    # 保存JSON文件
    print(f"正在保存到 {output_path}...")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)
    
    print(f"转换完成！共 {len(unique_items)} 条记录")

def merge_categories(douyin_file, multi_file, output_file, filter_cost_over_100=False, sort_by_cost_desc=False):
    """合并抖音图文和多品多图两个类别到一个JSON文件
    
    Args:
        douyin_file: 抖音图文数据文件路径
        multi_file: 多品多图数据文件路径
        output_file: 输出文件路径
        filter_cost_over_100: 是否过滤掉cost>100的记录（仅用于消耗最低数据）
        sort_by_cost_desc: 是否按cost降序排序（True=降序，False=升序，None=不排序）
    """
    categories = []
    
    # 读取抖音图文数据
    if douyin_file.exists():
        with open(douyin_file, 'r', encoding='utf-8') as f:
            douyin_data = json.load(f)
            if douyin_data.get('categories'):
                for cat in douyin_data['categories']:
                    # 如果需要过滤cost>100的记录
                    if filter_cost_over_100:
                        original_count = len(cat['items'])
                        cat['items'] = [
                            item for item in cat['items']
                            if item.get('cost', 0.0) <= 100.0
                        ]
                        filtered_count = original_count - len(cat['items'])
                        if filtered_count > 0:
                            print(f"    过滤掉 {filtered_count} 条cost>100的记录")
                    
                    # 按cost排序
                    if sort_by_cost_desc is not None:
                        cat['items'].sort(key=lambda x: x.get('cost', 0.0), reverse=sort_by_cost_desc)
                    
                    categories.append(cat)
    
    # 读取多品多图数据
    if multi_file.exists():
        with open(multi_file, 'r', encoding='utf-8') as f:
            multi_data = json.load(f)
            if multi_data.get('categories'):
                for cat in multi_data['categories']:
                    # 过滤掉不满足9张图的素材组（仅针对多品多图）
                    original_count = len(cat['items'])
                    cat['items'] = [
                        item for item in cat['items']
                        if len(item.get('material_urls', [])) == 9
                    ]
                    filtered_count = original_count - len(cat['items'])
                    if filtered_count > 0:
                        print(f"    过滤掉 {filtered_count} 条不满足9张图的记录")
                    
                    # 如果需要过滤cost>100的记录
                    if filter_cost_over_100:
                        original_count = len(cat['items'])
                        cat['items'] = [
                            item for item in cat['items']
                            if item.get('cost', 0.0) <= 100.0
                        ]
                        filtered_count = original_count - len(cat['items'])
                        if filtered_count > 0:
                            print(f"    过滤掉 {filtered_count} 条cost>100的记录")
                    
                    # 按cost排序
                    if sort_by_cost_desc is not None:
                        cat['items'].sort(key=lambda x: x.get('cost', 0.0), reverse=sort_by_cost_desc)
                    
                    categories.append(cat)
    
    # 合并后的数据结构
    merged_data = {
        "categories": categories
    }
    
    # 保存合并后的文件
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(merged_data, f, ensure_ascii=False, indent=2)
    
    print(f"合并完成: {output_file}")
    print(f"  包含 {len(categories)} 个类别")
    for cat in categories:
        print(f"    - {cat['name']}: {len(cat['items'])} 条记录")

if __name__ == "__main__":
    base_dir = Path(__file__).parent
    public_dir = base_dir / "public"
    
    # 转换多品多图消耗最高（尝试多个可能的文件名）
    excel_high = None
    for filename in ["多品多图消耗最高.xlsx", "多品多图消耗最高(1).xlsx"]:
        path = base_dir / filename
        if path.exists():
            excel_high = path
            break
    
    json_multi_high = public_dir / "data-多品多图消耗最高-temp.json"
    if excel_high and excel_high.exists():
        convert_excel_to_json(excel_high, json_multi_high, "多品多图消耗最高")
    else:
        print(f"文件不存在: 多品多图消耗最高.xlsx 或 多品多图消耗最高(1).xlsx")
    
    # 转换多品多图消耗最低
    excel_low = base_dir / "多品多图消耗最低.xlsx"
    json_multi_low = public_dir / "data-多品多图消耗最低-temp.json"
    if excel_low.exists():
        convert_excel_to_json(excel_low, json_multi_low, "多品多图消耗最低")
    else:
        print(f"文件不存在: {excel_low}")
    
    # 转换抖音图文消耗最高（支持xlsx和csv格式，包括带(1)的文件名）
    excel_douyin_high = None
    for filename in ["抖音图文消耗最高.xlsx", "抖音图文消耗最高.csv", "抖音图文消耗最高(1).csv", "抖音图文消耗最高(1).xlsx"]:
        path = base_dir / filename
        if path.exists():
            excel_douyin_high = path
            break
    
    json_douyin_high = public_dir / "data-抖音图文消耗最高-temp.json"
    if excel_douyin_high and excel_douyin_high.exists():
        convert_excel_to_json(excel_douyin_high, json_douyin_high, "抖音图文消耗最高", category_id='douyin', layout='strip')
    else:
        print(f"文件不存在: 抖音图文消耗最高.xlsx 或 抖音图文消耗最高.csv")
    
    # 转换抖音图文消耗最低
    excel_douyin_low = base_dir / "抖音图文消耗最低.xlsx"
    json_douyin_low = public_dir / "data-抖音图文消耗最低-temp.json"
    if excel_douyin_low.exists():
        convert_excel_to_json(excel_douyin_low, json_douyin_low, "抖音图文消耗最低", category_id='douyin', layout='strip')
    else:
        print(f"文件不存在: {excel_douyin_low}")
    
    # 合并消耗最高的数据（按cost降序排序）
    print("\n合并消耗最高的数据（按cost降序排序）...")
    merge_categories(
        json_douyin_high,
        json_multi_high,
        public_dir / "data-消耗最高.json",
        filter_cost_over_100=False,
        sort_by_cost_desc=True
    )
    
    # 合并消耗最低的数据（过滤掉cost>100的记录，按cost升序排序）
    print("\n合并消耗最低的数据（过滤cost>100的记录，按cost升序排序）...")
    merge_categories(
        json_douyin_low,
        json_multi_low,
        public_dir / "data-消耗最低.json",
        filter_cost_over_100=True,
        sort_by_cost_desc=False
    )
    
    # 清理临时文件
    print("\n清理临时文件...")
    for temp_file in [json_multi_high, json_multi_low, json_douyin_high, json_douyin_low]:
        if temp_file.exists():
            temp_file.unlink()
            print(f"  删除: {temp_file.name}")

