import os
import sys
import json
import time
import logging
from datetime import datetime
from collections import Counter

from excel_processor import read_excel
from scraper import scrape_url, get_analysis_text, detect_platform
from sentiment_analyzer import analyze_sentiment, classify_negative_reviews, generate_negative_summary

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def print_color(text, color='reset'):
    colors = {
        'red': '\033[91m', 'green': '\033[92m', 'yellow': '\033[93m',
        'blue': '\033[94m', 'cyan': '\033[96m', 'bold': '\033[1m', 'reset': '\033[0m',
    }
    output = text
    if sys.stdout.isatty():
        try:
            print(f"{colors.get(color, colors['reset'])}{output}{colors['reset']}")
        except UnicodeEncodeError:
            print(output.encode('ascii', errors='replace').decode('ascii'))
    else:
        try:
            print(output)
        except UnicodeEncodeError:
            print(output.encode('ascii', errors='replace').decode('ascii'))


def deduplicate_reviews(reviews: list[dict]) -> list[dict]:
    seen_links = set()
    deduped = []
    dup_count = 0
    for r in reviews:
        link = r['link'].strip().rstrip('/')
        if link in seen_links:
            dup_count += 1
            continue
        seen_links.add(link)
        r['link'] = link
        deduped.append(r)
    return deduped, dup_count


def save_progress(filepath: str, analyzed: list[dict], total: int):
    temp = filepath + '.progress.json'
    with open(temp, 'w', encoding='utf-8') as f:
        json.dump({'analyzed': len(analyzed), 'total': total, 'reviews': analyzed}, f, ensure_ascii=False)


def load_progress(filepath: str) -> tuple:
    prog_file = filepath + '.progress.json'
    if os.path.exists(prog_file):
        try:
            with open(prog_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data.get('analyzed', 0), data.get('reviews', [])
        except Exception:
            return 0, []
    return 0, []


def process_excel(file_path: str, output_dir: str = None):
    file_path = file_path.strip().strip('"').strip("'")
    if not os.path.exists(file_path):
        print_color(f'文件不存在: {file_path}', 'red')
        return

    print_color(f'\n{"="*60}', 'cyan')
    print_color('  多平台舆论分析工具 v2.0（优化版）', 'bold')
    print_color(f'  文件: {os.path.basename(file_path)}', 'cyan')
    print_color(f'{"="*60}\n', 'cyan')

    print_color('正在读取所有Sheet...', 'yellow')
    result = read_excel(file_path)

    if not result['success']:
        print_color(f'读取失败: {result.get("error", "未知错误")}', 'red')
        return

    if 'sheets' in result:
        print_color(f'发现 {len(result["sheets"])} 个Sheet:', 'green')
        for s in result['sheets']:
            cols = s.get('columns', {})
            print(f'  [{s["name"]}] {s["rows"]} 行, {s["valid"]} 条有效 '
                  f'(摘要: {cols.get("摘要","?")}, 链接: {cols.get("链接","?")})')

    reviews = result['reviews']
    print_color(f'\n总共读取 {len(reviews)} 条评价', 'green')

    if not reviews:
        print_color('未找到有效数据行', 'yellow')
        return

    if result.get('errors'):
        print_color(f'  跳过 {len(result["errors"])} 行', 'yellow')
        for err in result['errors'][:5]:
            print(f'    {err}')

    reviews, dup_count = deduplicate_reviews(reviews)
    print_color(f'去除重复链接: {dup_count} 条', 'yellow')
    print_color(f'待分析: {len(reviews)} 条评价\n', 'green')

    analyzed_start, analyzed_reviews = load_progress(file_path)
    if analyzed_start > 0:
        print_color(f'发现上次中断的进度，已分析 {analyzed_start}/{len(reviews)} 条', 'yellow')
        if input('是否继续上次进度? (y/n, 默认y): ').strip().lower() == 'n':
            analyzed_reviews = []
            analyzed_start = 0

    print_color(f'\n{"─"*60}', 'cyan')
    print_color('  开始逐条分析...', 'bold')

    normal_count = sum(1 for r in reviews if '正常' in r.get('sheet', ''))
    scrape_count = len(reviews) - normal_count
    if normal_count > 0:
        print_color(f'  发现 {normal_count} 条来自正常Sheet的评价，将直接使用摘要分析（跳过爬取）', 'yellow')
        print_color(f'  {scrape_count} 条来自敏感Sheet的评价将进行爬取分析', 'yellow')
    else:
        print_color(f'  共 {len(reviews)} 条，预计耗时较长，请耐心等待', 'yellow')
    print_color(f'{"─"*60}\n', 'cyan')

    start_time = time.time()
    for i, review in enumerate(reviews, 1):
        if i <= analyzed_start:
            continue

        link = review['link']
        short_link = link[:60] + '...' if len(link) > 60 else link
        elapsed = time.time() - start_time
        avg_time = elapsed / max(1, i - 1)
        remaining = avg_time * (len(reviews) - i + 1)
        is_normal_sheet = '正常' in review.get('sheet', '')

        if is_normal_sheet:
            print_color(f'  [{i}/{len(reviews)}] ({int(remaining/60)}分剩余) [摘要模式]', 'cyan')
        else:
            print_color(f'  [{i}/{len(reviews)}] ({int(remaining/60)}分剩余) {short_link}', 'yellow')
        sys.stdout.flush()

        if is_normal_sheet:
            scrape_result = {'scrape_success': False, 'platform': '', 'title': '', 'error': ''}
            analysis_text = review.get('summary', '')
        else:
            scrape_result = scrape_url(link, review.get('summary', ''))
            analysis_text = get_analysis_text(scrape_result, review.get('summary', ''))

        sentiment = analyze_sentiment(analysis_text)
        platform = review.get('platform') or scrape_result.get('platform') or detect_platform(link)

        analyzed = {
            'row_index': review.get('row_index', i),
            'summary': review.get('summary', ''),
            'link': link,
            'platform': platform,
            'sheet': review.get('sheet', ''),
            'title': scrape_result.get('title', ''),
            'scrape_success': scrape_result.get('scrape_success', False),
            'scrape_error': scrape_result.get('error', ''),
            'is_negative': sentiment['is_negative'],
            'sentiment_score': sentiment['score'],
            'negative_score': sentiment['negative_score'],
            'negative_keywords': sentiment['negative_keywords'],
            'positive_keywords': sentiment.get('positive_keywords', []),
            'confidence': sentiment.get('confidence', 'neutral'),
            'raw_negative_hits': sentiment.get('raw_negative_hits', 0),
            'raw_positive_hits': sentiment.get('raw_positive_hits', 0),
        }
        analyzed_reviews.append(analyzed)

        confidence_tag = {
            'high': '[高可信]', 'medium': '[中可信]', 'low': '[低可信]'
        }.get(sentiment.get('confidence', ''), '')

        if sentiment['is_negative']:
            print_color(f'    [-] 负面{confidence_tag} (负面分: {sentiment["negative_score"]:.2f})', 'red')
        else:
            pos_hits = sentiment.get('raw_positive_hits', 0)
            if pos_hits >= 2:
                print_color(f'    [+] 正面评价', 'green')
            else:
                print_color(f'    [+] 正常', 'green')

        if i % 100 == 0:
            save_progress(file_path, analyzed_reviews, len(reviews))

    save_progress(file_path, analyzed_reviews, len(reviews))

    total = len(analyzed_reviews)
    negative_reviews = classify_negative_reviews(analyzed_reviews)
    negative_count = len(negative_reviews)
    negative_pct = negative_count / total * 100 if total > 0 else 0

    summary_text = generate_negative_summary(negative_reviews)

    output_dir = output_dir or os.path.dirname(file_path)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_file = os.path.join(output_dir, f'{base_name}_分析报告_{timestamp}.txt')

    platform_stats = Counter(r['platform'] for r in analyzed_reviews)
    negative_platform_stats = Counter(r['platform'] for r in negative_reviews)

    confidence_levels = Counter(r.get('confidence', 'neutral') for r in negative_reviews)
    positive_keywords_all = []
    for r in analyzed_reviews:
        positive_keywords_all.extend(r.get('positive_keywords', []))

    report_lines = []
    report_lines.append(f'{"="*60}')
    report_lines.append(f'  多平台舆论分析报告 v2.0')
    report_lines.append(f'  文件: {os.path.basename(file_path)}')
    report_lines.append(f'  分析时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    report_lines.append(f'  总数据: {result.get("total_rows", total)} 行 | 去重后: {total} 条')
    report_lines.append(f'{"="*60}')
    report_lines.append('')
    report_lines.append(f'[分析概览]')
    report_lines.append(f'  - 总评价数: {total}')
    report_lines.append(f'  - 负面评价: {negative_count} ({negative_pct:.1f}%)')
    report_lines.append(f'     - 高可信: {confidence_levels.get("high", 0)} 条')
    report_lines.append(f'     - 中可信: {confidence_levels.get("medium", 0)} 条')
    report_lines.append(f'     - 低可信: {confidence_levels.get("low", 0)} 条')
    report_lines.append(f'  - 正常评价: {total - negative_count}')
    report_lines.append('')
    report_lines.append(f'[各平台分布]')
    for platform, count in platform_stats.most_common():
        neg_count = negative_platform_stats.get(platform, 0)
        bar_len = int(count / max(platform_stats.values()) * 20) if max(platform_stats.values()) > 0 else 0
        bar = '#' * bar_len
        report_lines.append(f'  {platform:　<8} {count:>4}条  负面{neg_count}条  {bar}')
    report_lines.append('')
    report_lines.append(f'[高频正面词]')
    if positive_keywords_all:
        top_pos = Counter(positive_keywords_all).most_common(10)
        report_lines.append('  ' + ', '.join(f'"{kw}"({c}次)' for kw, c in top_pos))
    report_lines.append('')
    report_lines.append(f'{"-"*60}')
    report_lines.append('')
    report_lines.append(summary_text)
    report_lines.append('')
    report_lines.append(f'{"-"*60}')
    report_lines.append('')
    report_lines.append(f'[全部评价详细结果]')
    report_lines.append('')
    for i, r in enumerate(analyzed_reviews, 1):
        tag = '[负面]' if r['is_negative'] else '[正常]'
        conf = {'high': '[高可信]', 'medium': '[中可信]', 'low': '[低可信]'}.get(r.get('confidence', ''), '')
        report_lines.append(f'  {i}. {tag}{conf} {r["platform"]}')
        report_lines.append(f'     摘要: {r["summary"][:120] if r["summary"] else "无摘要"}')
        report_lines.append(f'     链接: {r["link"]}')
        if r['is_negative']:
            report_lines.append(f'     负面指数: {r["negative_score"]:.2f} | 可信度: {r.get("confidence", "?")}')
            if r['negative_keywords']:
                kws = [k for k in r['negative_keywords'] if not k.startswith('[模式]')][:5]
                if kws:
                    report_lines.append(f'     匹配关键词: {", ".join(kws)}')
        report_lines.append('')

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report_lines))

    prog_file = file_path + '.progress.json'
    if os.path.exists(prog_file):
        os.remove(prog_file)

    print_color(f'\n{"="*60}', 'cyan')
    print_color(f'  [分析完成!]', 'bold')
    print_color(f'{"="*60}', 'cyan')
    print()
    print_color(f'  总评价: {total}', 'cyan')
    print_color(f'  正面关键词最多: {Counter(positive_keywords_all).most_common(1)[0][0] if positive_keywords_all else "无"}', 'green')
    print_color(f'  负面评价: {negative_count} 条 ({negative_pct:.1f}%)', 'red' if negative_count > 0 else 'green')
    print_color(f'    高可信: {confidence_levels.get("high", 0)} | 中可信: {confidence_levels.get("medium", 0)} | 低可信: {confidence_levels.get("low", 0)}', 'yellow')
    print_color(f'  正常评价: {total - negative_count} 条', 'green')
    print()

    if negative_reviews:
        print_color(f'  [负面评价列表:]', 'red')
        for i, r in enumerate(negative_reviews[:20], 1):
            plat = r['platform']
            summary = r['summary'][:60] if r['summary'] else '无摘要'
            conf_tag = {'high': '[高可信]', 'medium': '[中可信]', 'low': '[低可信]'}.get(r.get('confidence', ''), '')
            print_color(f'    {i}. {conf_tag}[{plat}] {summary}', 'red')
            print_color(f'       {r["link"]}', 'yellow')
        if len(negative_reviews) > 20:
            print_color(f'    ... 还有 {len(negative_reviews) - 20} 条负面评价', 'yellow')
        print()

    print_color(f'  报告已保存至:', 'green')
    print_color(f'  {output_file}', 'cyan')
    print()


def main():
    if len(sys.argv) > 1:
        process_excel(sys.argv[1])
    else:
        print_color('多平台舆论分析工具 v2.0（优化版）', 'bold')
        print_color('=' * 40, 'cyan')
        print_color('支持多Sheet读取 + 自动去重 + 中断续传', 'green')
        print()
        file_path = input('\n请拖入或输入 Excel 文件路径: ').strip()
        if file_path:
            process_excel(file_path)
        else:
            print_color('未输入文件路径', 'yellow')


if __name__ == '__main__':
    main()
