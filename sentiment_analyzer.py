import re
import logging
from snownlp import SnowNLP

logger = logging.getLogger(__name__)

# === 1. 负面关键词（按领域分类）===

# 美妆/护肤品专用负面词
BEAUTY_NEGATIVE_KEYWORDS = [
    '卡粉', '浮粉', '假白', '假面', '厚重', '油腻', '闷痘', '闭口',
    '过敏', '泛红', '刺痛', '拔干', '起皮', '搓泥', '不服帖',
    '脱妆', '斑驳', '暗沉', '氧化快', '遮瑕差', '持妆短', '持妆差',
    '刺激', '熏眼睛', '辣眼睛', '味道难闻', '香精重', '酒精味',
    '涂不匀', '抹不开', '推不开', '结块', '掉渣',
    '无功无过', '没效果', '没卵用', '鸡肋',
]

# 通用负面词（排除已在美妆词中的）
GENERAL_NEGATIVE_KEYWORDS = [
    '差评', '垃圾', '坑人', '骗人', '虚假宣传', '质量问题',
    '后悔', '上当', '不值', '太差', '很差', '不好', '不行', '糟糕',
    '恶心', '烂', '破', '假货', '黑心', '无良', '欺诈', '套路',
    '毛病', '缺陷', '难用', '失望', '暴利', '割韭菜', '智商税',
    '态度恶劣', '售后差', '推诿', '扯皮', '不处理',
    '踩雷',
    '偷工减料', '缩水', '加价',
]

# 强负面短语模式（每个模式带权重）
STRONG_NEGATIVE_PATTERNS = [
    (r'(千万别|千万不要|绝对不要|一定不要)', 4.0),
    (r'(避雷|避坑|踩雷|翻车)', 3.5),
    (r'(后悔|心碎|崩溃|愤怒|生气|心寒)', 3.0),
    (r'(维权|投诉|曝光|举报|起诉|报警)', 3.0),
    (r'(太差|太烂|太垃圾|太坑)', 3.0),
    (r'(不推荐|别买|不要买|不建议)', 3.0),
    (r'(割韭菜|智商税|交税)', 3.5),
    (r'(过敏|烂脸|毁脸)', 3.5),
    (r'(假货|售假|卖假)', 3.5),
]

# === 2. 正面关键词（用于抵消误判）===

POSITIVE_KEYWORDS = [
    '好用', '回购', '推荐', '轻薄', '自然', '服帖', '水润',
    '遮瑕好', '持妆久', '遮瑕力强', '滋润', '清爽', '透气',
    '细腻', '精致', '高级', '回购', '无限回购', '囤货',
    '性价比高', '平价', '物美价廉', '物超所值', '惊喜',
    '会回购', '已回购', '买了不亏', '值得买', '值得入手',
    '效果好', '效果明显', '皮肤变好', '改善', '透亮',
    '不卡粉', '不浮粉', '不假白', '不闷痘', '不刺激',
    '不油腻', '不厚重', '不拔干', '不搓泥',
]

# === 5. 营销/促销上下文排除模式 ===

PROMOTIONAL_CONTEXTS = [
    # 千万别 + 正面营销词
    (r'千万别\s*(错过|停产|不下架)', '千万别在营销语境中表示促销'),
    (r'求求.*千万别停产', '求别停产是正面评价'),
    # 举报 + 价格惊喜（假举报真促销）
    (r'举报.*(?:价格打到|到手这个价)', '举报+低价是营销手段'),
    (r'举报.*竟然.*价', '举报+竟然价是营销手段'),
    # 别买贵 = 别买贵的，是营销不是负面
    (r'别买贵', '别买贵表示别买贵的，促销语境'),
]

# === 6. 正面营销保证词（抵消强负面模式）===

POSITIVE_GUARANTEE_PHRASES = [
    '过敏包退', '过敏退', '过敏退款',
    '放心入', '安心入', '闭眼入',
    '福利', '优惠', '薅羊毛',
    '买一送一', '限时', '折扣',
]


def _is_promotional_marketing(text: str, matched_pattern: str, match_start: int, match_end: int) -> bool:
    before = text[max(0, match_start - 15):match_start]
    after = text[match_end:min(len(text), match_end + 30)]

    full_context = before + matched_pattern + after

    for promo_pattern, _ in PROMOTIONAL_CONTEXTS:
        if re.search(promo_pattern, full_context):
            return True

    if '千万别' in matched_pattern or '不要' in matched_pattern:
        if any(p in after for p in ['错过', '停产', '不下架']):
            return True

    if '举报' in matched_pattern:
        if any(p in after for p in ['价格打到', '到手这个价', '竟然']):
            return True

    # 检查是否在保证性短语附近（如"过敏包退"）
    for phrase in POSITIVE_GUARANTEE_PHRASES:
        if phrase in full_context:
            return True

    return False

NEGATION_WORDS = ['不', '没', '别', '不要', '不是', '没有', '不会', '不能', '并非']


def _has_negation_prefix(text: str, keyword: str) -> bool:
    idx = text.find(keyword)
    if idx < 0 or idx < 1:
        return False
    prev_char = text[idx - 1]
    if prev_char in ('不',):
        return True
    if idx >= 2 and text[idx - 2:idx] in ('不要', '不是', '没有', '不会', '不能', '并非', '别想'):
        return True
    return False


def analyze_sentiment(text: str) -> dict:
    if not text or not isinstance(text, str) or not text.strip():
        return {
            'is_negative': False,
            'score': 0.5,
            'negative_keywords': [],
            'negative_score': 0.0,
            'confidence': 'neutral',
        }

    found_negative = []
    found_positive = []
    negative_score = 0.0
    positive_score = 0.0

    for keyword in BEAUTY_NEGATIVE_KEYWORDS:
        if keyword in text:
            if _has_negation_prefix(text, keyword):
                found_positive.append(f'(非){keyword}')
                positive_score += 1.0
            else:
                found_negative.append(keyword)
                negative_score += 1.0

    for keyword in GENERAL_NEGATIVE_KEYWORDS:
        if keyword in text:
            if _has_negation_prefix(text, keyword):
                found_positive.append(f'(非){keyword}')
                positive_score += 1.0
            else:
                found_negative.append(keyword)
                negative_score += 1.0

    for pattern, weight in STRONG_NEGATIVE_PATTERNS:
        for match in re.finditer(pattern, text):
            matched_text = match.group()
            if _is_promotional_marketing(text, matched_text, match.start(), match.end()):
                found_positive.append(f'[营销排除]{pattern}')
                positive_score += 1.0
                continue
            found_negative.append(f'[模式]{pattern}')
            negative_score += weight

    for keyword in POSITIVE_KEYWORDS:
        if keyword in text:
            found_positive.append(keyword)
            positive_score += 1.0

    for phrase in POSITIVE_GUARANTEE_PHRASES:
        if phrase in text:
            found_positive.append(f'[保证]{phrase}')
            positive_score += 1.0

    raw_negative = negative_score
    raw_positive = positive_score

    negative_score = min(negative_score / 4.0, 1.0)
    positive_score = min(positive_score / 3.0, 1.0)

    try:
        snownlp_score = SnowNLP(text).sentiments
    except Exception:
        snownlp_score = 0.5

    keyword_net = negative_score * 0.7 - positive_score * 0.3
    combined_score = snownlp_score * 0.35 + (1.0 - max(0, keyword_net)) * 0.65
    combined_score = max(0.0, min(1.0, combined_score))

    is_negative = False
    confidence = 'neutral'

    if raw_negative >= 4.0 and negative_score > positive_score * 1.5:
        is_negative = True
        confidence = 'high'
    elif raw_negative >= 2.0 and negative_score > positive_score * 1.2:
        is_negative = True
        confidence = 'medium'
    elif combined_score < 0.35 and negative_score > positive_score:
        is_negative = True
        confidence = 'low'
    else:
        is_negative = False
        confidence = 'neutral' if combined_score >= 0.35 else 'positive'

    return {
        'is_negative': is_negative,
        'score': round(combined_score, 4),
        'negative_keywords': found_negative[:8],
        'positive_keywords': found_positive[:8],
        'negative_score': round(negative_score, 4),
        'positive_score': round(positive_score, 4),
        'confidence': confidence,
        'raw_negative_hits': raw_negative,
        'raw_positive_hits': raw_positive,
    }


def classify_negative_reviews(reviews: list[dict]) -> list[dict]:
    return [r for r in reviews if r.get('is_negative')]


def generate_negative_summary(negative_reviews: list[dict]) -> str:
    if not negative_reviews:
        return "未检测到明显的负面舆论评价。"

    total = len(negative_reviews)
    platforms = {}
    all_keywords = []

    for review in negative_reviews:
        platform = review.get('platform', '未知平台')
        platforms.setdefault(platform, []).append(review)
        all_keywords.extend(review.get('negative_keywords', []))

    platform_summary = '\n'.join(
        f"  - **{plat}**: {len(revs)} 条负面评价"
        for plat, revs in sorted(platforms.items(), key=lambda x: -len(x[1]))
    )

    top_keywords = []
    if all_keywords:
        from collections import Counter
        kw_counts = Counter(all_keywords)
        top_keywords = [f'"{kw}"({count}次)' for kw, count in kw_counts.most_common(10)]

    severity_levels = Counter(r.get('confidence', 'neutral') for r in negative_reviews)

    parts = [
        "## 负面舆论总结报告\n",
        "### 概览\n",
        f"共检测到 **{total}** 条负面评价，涉及 **{len(platforms)}** 个平台。\n",
        f"可信度分布：高置信度 {severity_levels.get('high', 0)} 条，"
        f"中置信度 {severity_levels.get('medium', 0)} 条，"
        f"低置信度 {severity_levels.get('low', 0)} 条。\n",
        "### 各平台分布\n",
        platform_summary,
    ]

    if top_keywords:
        parts.extend([
            "\n### 高频负面关键词\n",
            ", ".join(top_keywords),
        ])

    parts.append("\n### 详细负面评价列表\n")
    for i, r in enumerate(negative_reviews, 1):
        confidence_tag = {'high': '[高可信]', 'medium': '[中可信]', 'low': '[低可信]'}.get(
            r.get('confidence', 'neutral'), ''
        )
        parts.append(
            f"{i}. {confidence_tag} **平台**: {r.get('platform', '未知')}\n"
            f"   **摘要**: {r.get('summary', '无摘要')}\n"
            f"   **链接**: {r.get('link', '无链接')}\n"
            f"   **负面得分**: {r.get('negative_score', 0):.2f}\n"
        )

    return '\n'.join(parts)
