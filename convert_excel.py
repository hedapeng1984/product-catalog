#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel数据转换为JSON脚本
将 苏州趣享2026年中秋礼品报价单.xlsx 转换为 products.json
"""

import pandas as pd
import json
import re
import sys

# 设置UTF-8输出
sys.stdout.reconfigure(encoding='utf-8')

def clean_value(val):
    """清理数据值"""
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        return val
    val = str(val).strip()
    val = re.sub(r'\s+', ' ', val)  # 合并多余空格
    return val if val else ""

def extract_category_from_directory(file_path):
    """从目录Sheet提取分类信息"""
    df = pd.read_excel(file_path, sheet_name='目录', header=None)

    categories = []
    # 找到第一行中的大类（月饼、生鲜大闸蟹、零食礼盒等）
    for i in range(len(df)):
        row = df.iloc[i]
        first_val = clean_value(row[0])
        if first_val and first_val != '序号' and not first_val.isdigit():
            # 检查该行是否有"点击跳转"链接
            has_link = False
            for j in range(1, min(5, len(row))):
                if '点击跳转' in str(row[j]):
                    has_link = True
                    break
            if has_link:
                categories.append({
                    "name": first_val,
                    "row": i
                })

    return categories

def read_all_sheets(file_path):
    """读取所有Sheet并转换为JSON"""
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names

    data = {
        "categories": [],
        "products": []
    }

    # 全局递增ID
    product_id = 0

    category_map = {
        '哈根达斯': '月饼',
        '元祖': '月饼',
        '美心': '月饼',
        '爱维尔': '月饼',
        '得月楼': '月饼',
        '甜星': '月饼',
        '花园饼屋': '月饼',
    }

    for sheet_name in sheets:
        if sheet_name == '目录':
            continue

        print(f"处理Sheet: {sheet_name}")

        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # 找到表头行（通常在第1行）
        header_row = None
        for i in range(len(df)):
            row = [clean_value(x) for x in df.iloc[i]]
            if '品名' in row or '产品名称' in row or '商品' in row:
                header_row = i
                break

        if header_row is None:
            header_row = 1

        # 获取表头
        headers = [clean_value(x) for x in df.iloc[header_row]]

        # 找到品名、零售价、集采价列的索引
        name_idx = next((i for i, h in enumerate(headers) if '品名' in h or '产品名称' in h or '商品' in h), 1)
        price_idx = next((i for i, h in enumerate(headers) if '零售价' in h), None)
        bulk_price_idx = next((i for i, h in enumerate(headers) if '集采价' in h or '团购价' in h or '批发价' in h), None)
        desc_idx = next((i for i, h in enumerate(headers) if '配置' in h or '说明' in h or '简介' in h), None)
        size_idx = next((i for i, h in enumerate(headers) if '尺寸' in h or '规格' in h), None)
        shelf_idx = next((i for i, h in enumerate(headers) if '保质期' in h or '有效期' in h), None)
        remark_idx = next((i for i, h in enumerate(headers) if '备注' in h), None)

        # 读取产品数据
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            name = clean_value(row[name_idx] if name_idx < len(row) else "")

            if not name:
                continue

            product_id += 1
            product = {
                "id": product_id,
                "category": category_map.get(sheet_name, sheet_name),
                "brand": sheet_name,
                "name": name,
                "description": clean_value(row[desc_idx]) if desc_idx and desc_idx < len(row) else "",
                "retailPrice": clean_value(row[price_idx]) if price_idx and price_idx < len(row) else "",
                "bulkPrice": clean_value(row[bulk_price_idx]) if bulk_price_idx and bulk_price_idx < len(row) else "",
                "size": clean_value(row[size_idx]) if size_idx and size_idx < len(row) else "",
                "shelfLife": clean_value(row[shelf_idx]) if shelf_idx and shelf_idx < len(row) else "",
                "remark": clean_value(row[remark_idx]) if remark_idx and remark_idx < len(row) else "",
                "image": ""
            }

            data["products"].append(product)

    # 构建分类列表
    category_set = set()
    for p in data["products"]:
        category_set.add(p["category"])

    data["categories"] = [{"name": c, "brands": []} for c in sorted(category_set)]

    # 为每个分类添加品牌
    for cat in data["categories"]:
        brands = set()
        for p in data["products"]:
            if p["category"] == cat["name"]:
                brands.add(p["brand"])
        cat["brands"] = sorted(list(brands))

    return data

def main():
    file_path = r'C:\Users\dapen\Downloads\苏州趣享2026年中秋礼品报价单.xlsx'
    output_path = r'C:\Users\dapen\Downloads\product-catalog\products.json'

    print("开始转换Excel数据...")

    data = read_all_sheets(file_path)

    print(f"转换完成: {len(data['categories'])} 个分类, {len(data['products'])} 个产品")

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"已保存到: {output_path}")

    # 输出统计信息
    print("\n=== 数据统计 ===")
    for cat in data["categories"]:
        count = sum(1 for p in data["products"] if p["category"] == cat["name"])
        print(f"  {cat['name']}: {count} 个产品, 品牌: {', '.join(cat['brands'])}")

if __name__ == "__main__":
    main()
