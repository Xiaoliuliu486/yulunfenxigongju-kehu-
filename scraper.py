import re
import time
import logging
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

REQUEST_TIMEOUT = 8
MAX_RETRIES = 1
DELAY_BETWEEN_REQUESTS = 0.5

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate',
}


def _get_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def detect_platform(url: str) -> str:
    domain = urlparse(url).netloc.lower()
    if any(k in domain for k in ['douyin', 'iesdouyin']):
        return '抖音'
    if any(k in domain for k in ['kuaishou', 'gifshow']):
        return '快手'
    if any(k in domain for k in ['xiaohongshu', 'xhslink']):
        return '小红书'
    if any(k in domain for k in ['toutiao', '头条']):
        return '今日头条'
    if any(k in domain for k in ['dongchedi', '懂车帝']):
        return '懂车帝'
    if 'weibo' in domain:
        return '微博'
    if 'bilibili' in domain or 'b23' in domain:
        return 'B站'
    if 'zhihu' in domain:
        return '知乎'
    if 'weixin' in domain or 'qq' in domain:
        return '微信'
    return '其他'


def _extract_meta_content(soup: BeautifulSoup, selectors: list[str]) -> str:
    for selector in selectors:
        tag = soup.find('meta', attrs={'name': selector}) or soup.find('meta', attrs={'property': selector})
        if tag and tag.get('content'):
            return tag['content'].strip()
    return ''


def _extract_page_text(soup: BeautifulSoup) -> str:
    for tag in soup(['script', 'style', 'nav', 'footer', 'header', 'aside']):
        tag.decompose()

    text_parts = []
    for tag in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'span', 'div', 'li']):
        text = tag.get_text(strip=True)
        if len(text) > 10:
            text_parts.append(text)

    full_text = '\n'.join(text_parts)
    lines = [line.strip() for line in full_text.split('\n') if line.strip()]
    return '\n'.join(lines)


def scrape_url(url: str, existing_summary: str = '') -> dict:
    result = {
        'url': url,
        'platform': detect_platform(url),
        'title': '',
        'description': '',
        'extracted_text': '',
        'scrape_success': False,
        'error': '',
    }

    for attempt in range(MAX_RETRIES):
        try:
            session = _get_session()
            resp = session.get(url, timeout=REQUEST_TIMEOUT, allow_redirects=True)
            resp.raise_for_status()

            content_type = resp.headers.get('Content-Type', '')
            if 'video' in content_type or 'application' in content_type:
                result['error'] = '链接直接指向视频文件，无法解析页面文本'
                break

            soup = BeautifulSoup(resp.text, 'lxml')

            result['title'] = (
                _extract_meta_content(soup, ['og:title', 'twitter:title', 'title'])
                or (soup.title.string.strip() if soup.title else '')
            )
            result['description'] = _extract_meta_content(soup, [
                'og:description', 'twitter:description', 'description'
            ])
            result['extracted_text'] = _extract_page_text(soup)
            result['scrape_success'] = True
            break

        except requests.Timeout:
            result['error'] = f'请求超时 (尝试 {attempt + 1}/{MAX_RETRIES})'
            time.sleep(DELAY_BETWEEN_REQUESTS)
        except requests.RequestException as e:
            result['error'] = f'请求失败: {str(e)[:100]}'
            time.sleep(DELAY_BETWEEN_REQUESTS)
        except Exception as e:
            result['error'] = f'解析失败: {str(e)[:100]}'
            break

    return result


def get_analysis_text(scrape_result: dict, existing_summary: str = '') -> str:
    text_parts = []

    if existing_summary and existing_summary.strip():
        text_parts.append(existing_summary.strip())

    if scrape_result.get('description'):
        text_parts.append(scrape_result['description'])

    if scrape_result.get('title'):
        text_parts.append(scrape_result['title'])

    if scrape_result.get('extracted_text'):
        extracted = scrape_result['extracted_text']
        if len(extracted) > 2000:
            extracted = extracted[:2000]
        text_parts.append(extracted)

    return '\n'.join(text_parts)
