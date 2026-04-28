import os
import sys
import json
import time
import asyncio
import logging
import uuid
import tempfile
from datetime import datetime
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import uvicorn

from excel_processor import read_excel
from scraper import scrape_url, get_analysis_text, detect_platform
from sentiment_analyzer import analyze_sentiment, generate_negative_summary

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

executor = ThreadPoolExecutor(max_workers=4)

TASKS = {}


def process_analysis(file_path: str, task_id: str):
    try:
        TASKS[task_id] = {'status': 'processing', 'progress': 0, 'total': 0, 'message': '正在读取Excel文件...'}

        result = read_excel(file_path)
        if not result['success']:
            TASKS[task_id] = {'status': 'error', 'message': f'读取失败: {result.get("error", "未知错误")}'}
            return

        reviews = result['reviews']
        if not reviews:
            TASKS[task_id] = {'status': 'error', 'message': '未找到有效数据'}
            return

        TASKS[task_id] = {'status': 'processing', 'progress': 0, 'total': len(reviews), 'message': f'读取完成，共{len(reviews)}条，正在去重...'}

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
        reviews = deduped

        sheets_info = []
        if 'sheets' in result:
            for s in result['sheets']:
                sheets_info.append({
                    'name': s['name'],
                    'rows': s['rows'],
                    'valid': s['valid'],
                })

        total = len(reviews)
        TASKS[task_id]['total'] = total
        analyzed = []
        start_time = time.time()

        for i, review in enumerate(reviews, 1):
            is_normal = '正常' in review.get('sheet', '')

            elapsed = time.time() - start_time
            avg_time = elapsed / max(1, i - 1)
            remaining = avg_time * (total - i + 1)

            TASKS[task_id]['progress'] = i
            TASKS[task_id]['message'] = f'正在分析第 {i}/{total} 条...'
            TASKS[task_id]['eta'] = round(remaining, 1)

            if is_normal:
                analysis_text = review.get('summary', '')
                sentiment = analyze_sentiment(analysis_text)
            else:
                scrape_result = scrape_url(review['link'], review.get('summary', ''))
                analysis_text = get_analysis_text(scrape_result, review.get('summary', ''))
                sentiment = analyze_sentiment(analysis_text)

            analyzed.append({
                'row_index': review.get('row_index', i),
                'summary': review.get('summary', ''),
                'link': review.get('link', ''),
                'platform': review.get('platform', detect_platform(review.get('link', ''))),
                'sheet': review.get('sheet', ''),
                'title': review.get('title', ''),
                'is_negative': sentiment['is_negative'],
                'sentiment_score': sentiment['score'],
                'confidence': sentiment['confidence'],
                'negative_keywords': sentiment.get('negative_keywords', []),
                'positive_keywords': sentiment.get('positive_keywords', []),
                'negative_score': sentiment['negative_score'],
            })

        negative_reviews = [r for r in analyzed if r['is_negative']]
        summary_report = generate_negative_summary(negative_reviews)

        keyword_counts = Counter()
        for r in negative_reviews:
            for kw in r.get('negative_keywords', []):
                keyword_counts[kw] += 1

        platform_dist = Counter(r.get('platform', '未知') for r in negative_reviews)
        conf_dist = Counter(r.get('confidence', 'neutral') for r in negative_reviews)

        TASKS[task_id] = {
            'status': 'completed',
            'total': total,
            'duplicates': dup_count,
            'negative_count': len(negative_reviews),
            'negative_rate': round(len(negative_reviews) / total * 100, 1) if total > 0 else 0,
            'summary_report': summary_report,
            'sheets_info': sheets_info,
            'statistics': {
                'total': total,
                'duplicates': dup_count,
                'valid': total,
                'negative': len(negative_reviews),
                'negative_rate': f'{round(len(negative_reviews) / total * 100, 1) if total > 0 else 0}%',
                'confidence_levels': {
                    'high': conf_dist.get('high', 0),
                    'medium': conf_dist.get('medium', 0),
                    'low': conf_dist.get('low', 0),
                },
                'platform_distribution': dict(platform_dist),
                'top_negative_keywords': {kw: count for kw, count in keyword_counts.most_common(15)},
            },
            'negative_reviews': [
                {
                    'summary': r['summary'][:120] if r['summary'] else '',
                    'link': r['link'],
                    'platform': r['platform'],
                    'confidence': r['confidence'],
                    'negative_score': r['negative_score'],
                    'keywords': r['negative_keywords'][:5],
                }
                for r in negative_reviews
            ],
            'message': f'分析完成！共发现 {len(negative_reviews)} 条负面评价',
        }

    except Exception as e:
        logger.exception(f'分析过程出错: {e}')
        TASKS[task_id] = {'status': 'error', 'message': f'分析失败: {str(e)}'}
    finally:
        try:
            os.remove(file_path)
        except:
            pass


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info('舆论分析API服务启动')
    yield
    logger.info('舆论分析API服务关闭')
    for task_id in list(TASKS.keys()):
        if TASKS[task_id].get('status') == 'processing':
            TASKS[task_id]['status'] = 'cancelled'


app = FastAPI(
    title='爱敬舆论分析API',
    description='多平台舆论分析工具 - 上传Excel，自动分析负面评价',
    version='2.0.0',
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_credentials=True,
    allow_methods=['*'],
    allow_headers=['*'],
)


@app.get('/')
async def root():
    return {
        'service': '爱敬舆论分析API',
        'version': '2.0.0',
        'status': 'running',
        'endpoints': {
            'POST /analyze': '上传Excel文件并启动分析',
            'GET  /status/{task_id}': '查询分析进度',
            'GET  /result/{task_id}': '获取分析结果',
            'GET  /health': '健康检查',
        },
    }


@app.get('/health')
async def health():
    active_tasks = sum(1 for t in TASKS.values() if t.get('status') == 'processing')
    return {'status': 'healthy', 'active_tasks': active_tasks, 'timestamp': datetime.now().isoformat()}


@app.post('/analyze')
async def analyze(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail='请上传 .xlsx 或 .xls 格式的Excel文件')

    task_id = str(uuid.uuid4())[:8]

    upload_dir = os.path.join(os.path.dirname(__file__), 'uploads')
    os.makedirs(upload_dir, exist_ok=True)
    file_path = os.path.join(upload_dir, f'{task_id}_{file.filename}')

    content = await file.read()
    with open(file_path, 'wb') as f:
        f.write(content)

    TASKS[task_id] = {'status': 'queued', 'message': '任务已创建，等待处理...'}

    loop = asyncio.get_event_loop()
    loop.run_in_executor(executor, process_analysis, file_path, task_id)

    return {
        'task_id': task_id,
        'status': 'queued',
        'message': '分析任务已提交，请使用task_id查询进度',
        'query_url': f'/status/{task_id}',
    }


@app.get('/status/{task_id}')
async def get_status(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail='任务不存在')

    if task['status'] in ('completed', 'error'):
        return task

    return {
        'status': task['status'],
        'progress': task.get('progress', 0),
        'total': task.get('total', 0),
        'message': task.get('message', ''),
        'eta': task.get('eta', 0),
    }


@app.get('/result/{task_id}')
async def get_result(task_id: str):
    task = TASKS.get(task_id)
    if not task:
        raise HTTPException(status_code=404, detail='任务不存在')
    if task['status'] == 'error':
        raise HTTPException(status_code=500, detail=task.get('message', '分析失败'))
    if task['status'] != 'completed':
        raise HTTPException(status_code=400, detail=f'任务尚未完成，当前状态: {task["status"]}')
    return task


@app.post('/analyze-sync')
async def analyze_sync(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail='请上传 .xlsx 或 .xls 格式的Excel文件')

    start_time = time.time()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        result = read_excel(tmp_path)
        if not result['success']:
            raise HTTPException(status_code=400, detail=f'读取失败: {result.get("error", "未知错误")}')

        reviews = result['reviews']
        if not reviews:
            raise HTTPException(status_code=400, detail='未找到有效数据')

        seen_links = set()
        deduped = []
        dup_count = 0
        for r in reviews:
            link = r['link'].strip().rstrip('/') if r.get('link') else ''
            if link in seen_links:
                dup_count += 1
                continue
            seen_links.add(link)
            r['link'] = link
            deduped.append(r)
        reviews = deduped

        max_reviews = min(len(reviews), 500)
        reviews = reviews[:max_reviews]
        total = len(reviews)

        analyzed = []
        for i, review in enumerate(reviews, 1):
            is_normal = '正常' in review.get('sheet', '')
            if is_normal:
                analysis_text = review.get('summary', '')
            else:
                analysis_text = review.get('summary', '')
            sentiment = analyze_sentiment(analysis_text)

            analyzed.append({
                'summary': (review.get('summary', '') or '')[:120],
                'link': review.get('link', ''),
                'platform': review.get('platform', '') or detect_platform(review.get('link', '')),
                'confidence': sentiment['confidence'],
                'is_negative': sentiment['is_negative'],
                'negative_score': sentiment['negative_score'],
                'keywords': sentiment.get('negative_keywords', [])[:5],
            })

        negative_reviews = [r for r in analyzed if r['is_negative']]
        summary_report = generate_negative_summary(negative_reviews)

        keyword_counts = Counter()
        for r in negative_reviews:
            for kw in r.get('keywords', []):
                keyword_counts[kw] += 1

        platform_dist = Counter(r.get('platform', '未知') for r in negative_reviews)
        conf_dist = Counter(r.get('confidence', 'neutral') for r in negative_reviews)

        elapsed = round(time.time() - start_time, 1)

        return {
            'status': 'completed',
            'total': total,
            'total_original': result.get('total', total),
            'duplicates': dup_count,
            'negative_count': len(negative_reviews),
            'negative_rate': f'{round(len(negative_reviews) / total * 100, 1) if total > 0 else 0}%',
            'processing_time': f'{elapsed}秒',
            'summary_report': summary_report,
            'statistics': {
                'total': total,
                'duplicates': dup_count,
                'negative': len(negative_reviews),
                'negative_rate': f'{round(len(negative_reviews) / total * 100, 1) if total > 0 else 0}%',
                'confidence_levels': {
                    'high': conf_dist.get('high', 0),
                    'medium': conf_dist.get('medium', 0),
                    'low': conf_dist.get('low', 0),
                },
                'platform_distribution': dict(platform_dist),
                'top_negative_keywords': {kw: count for kw, count in keyword_counts.most_common(15)},
            },
            'negative_reviews': negative_reviews,
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f'同步分析出错: {e}')
        raise HTTPException(status_code=500, detail=f'分析失败: {str(e)}')
    finally:
        try:
            os.unlink(tmp_path)
        except:
            pass


if __name__ == '__main__':
    port = int(os.environ.get('PORT', sys.argv[1] if len(sys.argv) > 1 else 8000))
    uvicorn.run(app, host='0.0.0.0', port=port)
