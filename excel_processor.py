import os
import logging
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

COLUMN_MAPPING = {
    '摘要': ['摘要', 'summary', '内容摘要', '简介', '简述', '评价内容', '评论', '内容', '文本', '标题/微博内容'],
    '链接': ['链接', 'link', 'url', '网址', '地址', 'URL', '原文/评论链接'],
    '平台': ['平台', 'platform', '来源', '来源网站', '渠道', '站点'],
    '标题': ['标题', 'title', '主题', '名称'],
}


def _find_column(df: pd.DataFrame, possible_names: list[str]) -> Optional[str]:
    for name in possible_names:
        if name in df.columns:
            return name
    for name in possible_names:
        matching = [col for col in df.columns if name.lower() in col.lower()]
        if matching:
            return matching[0]
    return None


def _parse_sheet(df: pd.DataFrame, sheet_name: str = '') -> dict:
    if df.empty:
        return {'reviews': [], 'errors': [], 'total_rows': 0}

    summary_col = _find_column(df, COLUMN_MAPPING['摘要'])
    link_col = _find_column(df, COLUMN_MAPPING['链接'])
    platform_col = _find_column(df, COLUMN_MAPPING['平台'])
    title_col = _find_column(df, COLUMN_MAPPING['标题'])

    if not link_col:
        return {
            'reviews': [],
            'errors': [f'[Sheet: {sheet_name}] 未找到链接列，列名: {list(df.columns)[:10]}...'],
            'total_rows': len(df),
        }

    reviews = []
    errors = []

    for idx, row in df.iterrows():
        summary = str(row.get(summary_col, '')) if summary_col and pd.notna(row.get(summary_col)) else ''
        link = str(row.get(link_col, '')) if link_col and pd.notna(row.get(link_col)) else ''
        platform = str(row.get(platform_col, '')) if platform_col and pd.notna(row.get(platform_col)) else ''
        title = str(row.get(title_col, '')) if title_col and pd.notna(row.get(title_col)) else ''

        if not link or link.lower() in ['nan', 'none', '']:
            errors.append(f'第 {idx + 2} 行缺少链接')
            continue

        review = {
            'row_index': idx + 2,
            'summary': summary if summary and summary != 'nan' else '',
            'link': link.strip(),
            'platform': platform if platform and platform != 'nan' else '',
            'title': title if title and title != 'nan' else '',
            'sheet': sheet_name,
        }
        reviews.append(review)

    return {
        'reviews': reviews,
        'errors': errors,
        'total_rows': len(df),
        'columns_found': {
            '摘要': summary_col or '未找到',
            '链接': link_col or '未找到',
            '平台': platform_col or '未找到',
            '标题': title_col or '未找到',
        },
    }


def read_excel(file_path: str, sheet_name: str = None) -> dict:
    if not os.path.exists(file_path):
        return {'success': False, 'error': f'文件不存在: {file_path}', 'reviews': []}

    try:
        if sheet_name:
            df = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_name)
            result = _parse_sheet(df, sheet_name)
            result['success'] = True
            result['error'] = ''
            result['valid_rows'] = len(result['reviews'])
            return result
        else:
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            all_reviews = []
            all_errors = []
            total_rows = 0
            sheets_info = []

            for name in xls.sheet_names:
                df = pd.read_excel(file_path, engine='openpyxl', sheet_name=name)
                total_rows += len(df)
                result = _parse_sheet(df, name)
                all_reviews.extend(result['reviews'])
                all_errors.extend(result['errors'])
                sheets_info.append({
                    'name': name,
                    'rows': len(df),
                    'valid': len(result['reviews']),
                    'columns': result.get('columns_found', {}),
                })

            return {
                'success': True,
                'error': '',
                'reviews': all_reviews,
                'total_rows': total_rows,
                'valid_rows': len(all_reviews),
                'errors': all_errors,
                'sheets': sheets_info,
            }
    except Exception as e:
        return {'success': False, 'error': f'读取Excel失败: {str(e)}', 'reviews': []}
