import csv
import json
import math
import os
import re
import sys
import time
import unicodedata
from statistics import median

from paddleocr import PaddleOCR


# ============================================================
# 配置
# ============================================================

IMAGE_PATH = "test.png"
OUTPUT_DIR = "output"

# OCR文字中心Y与节次中心Y最大匹配距离
# 实际值会根据课程表行距自动计算
PERIOD_MATCH_RATIO = 0.45

# 低置信度阈值
LOW_CONFIDENCE_THRESHOLD = 0.90


# ============================================================
# 明确不是课程的文字
#
# 注意：
# “班会”“劳动”“体育活动”等不能过滤，
# 因为它们可能是真正的课程。
# ============================================================

NON_COURSE_KEYWORDS = [
    "备注",
    "说明",
    "课程安排说明",
    "每学期安排",
    "每周安排",
]


NON_COURSE_EXACT = {
    "课间操",
    "大课间",
    "午休",
    "午睡",
    "眼保健操",
    "广播操",
    "早操",
    "升旗",
    "做操",

    # 时间段/非课程标识
    "上午",
    "下午",
    "中午",
    "早读",
    "晨读",
}


# ============================================================
# OCR文字基础处理
# ============================================================

def normalize_text(text):
    if text is None:
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(text)
    )

    text = text.strip()

    text = re.sub(
        r"[ \t]+",
        " ",
        text
    )

    return text


def box_to_list(box):
    if hasattr(box, "tolist"):
        box = box.tolist()

    return [
        float(x)
        for x in box
    ]


def get_center(box):
    x1, y1, x2, y2 = box

    return (
        (x1 + x2) / 2,
        (y1 + y2) / 2
    )


# ============================================================
# 中文数字
# ============================================================

CHINESE_DIGITS = {
    "零": 0,
    "〇": 0,
    "一": 1,
    "二": 2,
    "两": 2,
    "三": 3,
    "四": 4,
    "五": 5,
    "六": 6,
    "七": 7,
    "八": 8,
    "九": 9,
    "十": 10,
}


def chinese_number_to_int(text):
    """
    支持：
        一
        二
        九
        十
        十一
        十二
        二十
    """

    text = normalize_text(text)

    if text.isdigit():
        return int(text)

    if text in CHINESE_DIGITS:
        return CHINESE_DIGITS[text]

    if "十" not in text:
        return None

    # 十
    if text == "十":
        return 10

    # 十一 ~ 十九
    if text.startswith("十"):
        tail = text[1:]

        if tail in CHINESE_DIGITS:
            return 10 + CHINESE_DIGITS[tail]

    # 二十
    if text.endswith("十"):
        head = text[:-1]

        if head in CHINESE_DIGITS:
            return CHINESE_DIGITS[head] * 10

    # 二十一 ~ 二十九
    m = re.match(
        r"^([一二三四五六七八九])十([一二三四五六七八九])$",
        text
    )

    if m:
        return (
            CHINESE_DIGITS[m.group(1)] * 10
            + CHINESE_DIGITS[m.group(2)]
        )

    return None


# ============================================================
# 节次解析
# ============================================================

def parse_period(text):
    """
    支持：

        1
        2
        第1节
        第2课
        第一节
        第二节
        第十一节
        第3-4节
        第一～二节

    返回：

    {
        "start": 1,
        "end": 1,
        "label": "1"
    }
    """

    text = normalize_text(text)

    # 只允许它是一个纯节次表达式
    pattern = re.compile(
        r"^"
        r"(?:第\s*)?"
        r"([0-9]{1,2}|[一二三四五六七八九十百两]+)"
        r"(?:\s*[-~～—－]\s*"
        r"([0-9]{1,2}|[一二三四五六七八九十百两]+))?"
        r"\s*(?:节|课)?"
        r"$"
    )

    m = pattern.match(text)

    if not m:
        return None

    start_raw = m.group(1)
    end_raw = m.group(2)

    start = chinese_number_to_int(
        start_raw
    )

    if start is None:
        return None

    if end_raw:
        end = chinese_number_to_int(
            end_raw
        )
    else:
        end = start

    if end is None:
        return None

    if end < start:
        return None

    if start <= 0 or start > 30:
        return None

    return {
        "start": start,
        "end": end,
        "label": (
            f"{start}-{end}"
            if end != start
            else str(start)
        )
    }


# ============================================================
# 星期解析
# ============================================================

def parse_explicit_weekday(text):

    text = normalize_text(text)

    mapping = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "日": 7,
        "天": 7,

        "1": 1,
        "2": 2,
        "3": 3,
        "4": 4,
        "5": 5,
        "6": 6,
        "7": 7,
    }

    m = re.match(
        r"^(?:星期|周)\s*"
        r"([一二三四五六日天1-7])"
        r"$",
        text
    )

    if m:
        return mapping.get(
            m.group(1)
        )

    english = {
        "mon": 1,
        "monday": 1,
        "tue": 2,
        "tues": 2,
        "tuesday": 2,
        "wed": 3,
        "wednesday": 3,
        "thu": 4,
        "thur": 4,
        "thurs": 4,
        "thursday": 4,
        "fri": 5,
        "friday": 5,
        "sat": 6,
        "saturday": 6,
        "sun": 7,
        "sunday": 7,
    }

    return english.get(
        text.lower()
    )


# ============================================================
# 是否是明显非课程文字
# ============================================================

def is_obvious_non_course(text):

    text = normalize_text(text)

    if not text:
        return True

    if text in NON_COURSE_EXACT:
        return True

    for keyword in NON_COURSE_KEYWORDS:
        if keyword in text:
            return True

    return False


# ============================================================
# 按Y聚类
# ============================================================

def cluster_by_y(items, tolerance):

    if not items:
        return []

    items = sorted(
        items,
        key=lambda x: x["cy"]
    )

    groups = []

    current = [
        items[0]
    ]

    for item in items[1:]:

        avg_y = median(
            x["cy"]
            for x in current
        )

        if abs(
            item["cy"] - avg_y
        ) <= tolerance:

            current.append(item)

        else:

            groups.append(
                current
            )

            current = [
                item
            ]

    groups.append(
        current
    )

    return groups


# ============================================================
# 1维KMeans
#
# 仅用于“图片没有星期表头”时的最后备用方案
# ============================================================

def kmeans_1d(values, k):

    if len(values) < k:
        return None

    values = sorted(
        float(v)
        for v in values
    )

    # 分位点初始化
    centers = []

    for i in range(k):
        pos = (
            i
            * (len(values) - 1)
            / (k - 1)
            if k > 1
            else 0
        )

        centers.append(
            values[
                int(round(pos))
            ]
        )

    for _ in range(30):

        clusters = [
            []
            for _ in range(k)
        ]

        for value in values:

            idx = min(
                range(k),
                key=lambda j:
                abs(
                    value
                    - centers[j]
                )
            )

            clusters[idx].append(
                value
            )

        new_centers = []

        for idx, cluster in enumerate(
            clusters
        ):

            if cluster:
                new_centers.append(
                    sum(cluster)
                    / len(cluster)
                )
            else:
                new_centers.append(
                    centers[idx]
                )

        if all(
            abs(
                new_centers[i]
                - centers[i]
            ) < 0.01
            for i in range(k)
        ):
            centers = new_centers
            break

        centers = new_centers

    pairs = sorted(
        zip(
            centers,
            clusters
        ),
        key=lambda x: x[0]
    )

    centers = [
        x[0]
        for x in pairs
    ]

    clusters = [
        x[1]
        for x in pairs
    ]

    return (
        centers,
        clusters
    )


# ============================================================
# 自动寻找星期列
# ============================================================

def detect_day_columns(items):

    max_y = max(
        item["cy"]
        for item in items
    )

    # --------------------------------------------------------
    # 方案1：明确的“星期一 / 周一 / Monday”
    # --------------------------------------------------------

    explicit = []

    for item in items:

        day = parse_explicit_weekday(
            item["text"]
        )

        if day is None:
            continue

        # 表头通常在图片上部
        if item["cy"] > max_y * 0.45:
            continue

        explicit.append({
            **item,
            "day_number": day
        })

    if len(explicit) >= 3:

        by_day = {}

        for item in explicit:

            day = item[
                "day_number"
            ]

            if (
                day not in by_day
                or item["score"]
                > by_day[day]["score"]
            ):
                by_day[day] = item

        if len(by_day) >= 3:

            result = [
                {
                    "day_number": day,
                    "label":
                        f"周{day_to_chinese(day)}",
                    "cx":
                        item["cx"],
                    "cy":
                        item["cy"],
                    "source":
                        item["text"]
                }
                for day, item
                in by_day.items()
            ]

            result.sort(
                key=lambda x:
                x["cx"]
            )

            return result


    # --------------------------------------------------------
    # 方案2：数字星期表头
    #
    # complex2：
    #
    # 1  2  3  4  5
    # --------------------------------------------------------

    numeric = []

    for item in items:

        text = normalize_text(
            item["text"]
        )

        if not re.fullmatch(
            r"[1-7]",
            text
        ):
            continue

        if item["cy"] > max_y * 0.45:
            continue

        numeric.append(item)

    groups = cluster_by_y(
        numeric,
        tolerance=20
    )

    candidates = []

    for group in groups:

        group = sorted(
            group,
            key=lambda x:
            x["cx"]
        )

        numbers = [
            int(
                x["text"]
            )
            for x in group
        ]

        # 去重
        unique_numbers = list(
            dict.fromkeys(
                numbers
            )
        )

        if len(unique_numbers) < 3:
            continue

        # 必须是连续的1、2、3...
        expected = list(
            range(
                1,
                len(unique_numbers)
                + 1
            )
        )

        if unique_numbers != expected:
            continue

        candidates.append(
            group
        )

    if candidates:

        best = max(
            candidates,
            key=lambda g:
            len(g)
        )

        result = []

        for item in best:

            number = int(
                item["text"]
            )

            result.append({
                "day_number":
                    number,
                "label":
                    f"周{day_to_chinese(number)}",
                "cx":
                    item["cx"],
                "cy":
                    item["cy"],
                "source":
                    item["text"]
            })

        return result


    # --------------------------------------------------------
    # 方案3：中文“一二三四五”
    # --------------------------------------------------------

    chinese_week_chars = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "日": 7,
    }

    chinese_items = []

    for item in items:

        text = normalize_text(
            item["text"]
        )

        if text not in chinese_week_chars:
            continue

        if item["cy"] > max_y * 0.45:
            continue

        chinese_items.append(item)

    groups = cluster_by_y(
        chinese_items,
        tolerance=20
    )

    candidates = []

    for group in groups:

        group = sorted(
            group,
            key=lambda x:
            x["cx"]
        )

        nums = [
            chinese_week_chars[
                x["text"]
            ]
            for x in group
        ]

        if len(nums) >= 3:

            if nums == list(
                range(
                    1,
                    len(nums) + 1
                )
            ):

                candidates.append(
                    group
                )

    if candidates:

        best = max(
            candidates,
            key=lambda g:
            len(g)
        )

        return [
            {
                "day_number":
                    chinese_week_chars[
                        x["text"]
                    ],
                "label":
                    f"周{day_to_chinese(chinese_week_chars[x['text']])}",
                "cx":
                    x["cx"],
                "cy":
                    x["cy"],
                "source":
                    x["text"]
            }
            for x in best
        ]


    # --------------------------------------------------------
    # 方案4：完全没有星期表头
    #
    # 用所有下方文字的X坐标估算列中心
    # --------------------------------------------------------

    candidate_items = []

    for item in items:

        if item["cy"] < max_y * 0.20:
            continue

        if parse_period(
            item["text"]
        ):
            continue

        text = item["text"]

        if not text:
            continue

        candidate_items.append(
            item
        )


    if len(candidate_items) >= 12:

        values = [
            x["cx"]
            for x in candidate_items
        ]

        candidates = []

        for k in range(3, 8):

            result = kmeans_1d(
                values,
                k
            )

            if result is None:
                continue

            centers, clusters = result

            if len(centers) < 3:
                continue

            gaps = [
                centers[i]
                - centers[i - 1]
                for i in range(
                    1,
                    len(centers)
                )
            ]

            if not gaps:
                continue

            avg_gap = sum(gaps) / len(gaps)

            if avg_gap <= 0:
                continue

            std = math.sqrt(
                sum(
                    (
                        gap
                        - avg_gap
                    ) ** 2
                    for gap in gaps
                )
                / len(gaps)
            )

            cv = (
                std
                / avg_gap
            )

            singleton_rate = (
                sum(
                    1
                    for cluster
                    in clusters
                    if len(cluster) <= 1
                )
                / len(clusters)
            )

            score = (
                cv
                + singleton_rate * 0.5
            )

            candidates.append(
                (
                    score,
                    centers
                )
            )

        if candidates:

            _, centers = min(
                candidates,
                key=lambda x:
                x[0]
            )

            return [
                {
                    "day_number":
                        i + 1,
                    "label":
                        f"周{day_to_chinese(i + 1)}",
                    "cx":
                        center,
                    "cy":
                        None,
                    "source":
                        "自动推断"
                }
                for i, center
                in enumerate(centers)
            ]

    return []


def day_to_chinese(number):

    mapping = {
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "日",
    }

    return mapping.get(
        number,
        str(number)
    )


# ============================================================
# 找到节次编号
# ============================================================

def detect_periods(
    items,
    day_columns
):

    if not day_columns:
        return []

    day_centers = [
        x["cx"]
        for x in day_columns
    ]

    day_centers.sort()

    if len(day_centers) >= 2:

        gaps = [
            day_centers[i]
            - day_centers[i - 1]
            for i in range(
                1,
                len(day_centers)
            )
        ]

        column_gap = median(
            gaps
        )

    else:

        column_gap = 180


    # --------------------------------------------------------
    # 左侧节次区域
    #
    # 使用 0.72 倍列间距。
    #
    # 这样：
    #
    # complex1：
    # 周一≈349
    # 第一节≈187
    #
    # complex2：
    # 周一≈337
    # 第一节≈180
    #
    # 都能够被包含。
    # --------------------------------------------------------

    grid_left = (
        day_centers[0]
        - column_gap * 0.72
    )


    max_y = max(
        item["cy"]
        for item in items
    )


    period_items = []

    for item in items:

        if (
            item["cx"]
            >= grid_left
        ):
            continue

        if item["cy"] < max_y * 0.15:
            continue

        parsed = parse_period(
            item["text"]
        )

        if parsed is None:
            continue

        period_items.append({
            **item,
            **parsed
        })


    if not period_items:
        return []


    # --------------------------------------------------------
    # 同一节次可能被OCR重复检测
    # 通过Y聚类去重
    # --------------------------------------------------------

    groups = cluster_by_y(
        period_items,
        tolerance=28
    )


    anchors = []

    for group in groups:

        best = max(
            group,
            key=lambda x:
            x["score"]
        )

        anchors.append(
            best
        )


    anchors.sort(
        key=lambda x:
        x["cy"]
    )


    # --------------------------------------------------------
    # 判断节次编号是否重复
    #
    # complex1：
    # 1 2 3 4 5 6 7
    #
    # complex2：
    # 1 2 3 4 1 2 3 4
    #
    # 第二种要转换成：
    # 1 2 3 4 5 6 7 8
    # --------------------------------------------------------

    starts = [
        x["start"]
        for x in anchors
    ]

    unique_starts = len(
        set(starts)
    ) == len(starts)

    increasing = all(
        starts[i]
        > starts[i - 1]
        for i in range(
            1,
            len(starts)
        )
    )

    use_raw_number = (
        unique_starts
        and increasing
    )


    periods = []

    current_number = 1

    for anchor in anchors:

        if use_raw_number:

            normalized_number = (
                anchor["start"]
            )

        else:

            normalized_number = (
                current_number
            )


        span = (
            anchor["end"]
            - anchor["start"]
            + 1
        )


        periods.append({
            "raw_label":
                anchor["label"],

            "raw_start":
                anchor["start"],

            "raw_end":
                anchor["end"],

            "number":
                normalized_number,

            "label":
                (
                    f"第{normalized_number}"
                    f"节"
                ),

            "cy":
                anchor["cy"],

            "cx":
                anchor["cx"],

            "score":
                anchor["score"]
        })


        if use_raw_number:

            current_number = (
                normalized_number
                + span
            )

        else:

            current_number += span


    return periods


# ============================================================
# 计算文字属于哪个节次
# ============================================================

def calculate_period_threshold(
    periods
):

    if len(periods) < 2:
        return 45

    ys = [
        p["cy"]
        for p in periods
    ]

    gaps = [
        ys[i]
        - ys[i - 1]
        for i in range(
            1,
            len(ys)
        )
    ]

    positive_gaps = [
        gap
        for gap in gaps
        if gap > 5
    ]

    if not positive_gaps:
        return 45

    typical_gap = median(
        positive_gaps
    )

    threshold = (
        typical_gap
        * PERIOD_MATCH_RATIO
    )

    # 不要太小，也不要太大
    threshold = max(
        30,
        threshold
    )

    threshold = min(
        60,
        threshold
    )

    return threshold


def assign_period(
    item,
    periods,
    threshold
):

    if not periods:
        return None, None

    nearest = min(
        periods,
        key=lambda p:
        abs(
            item["cy"]
            - p["cy"]
        )
    )

    distance = abs(
        item["cy"]
        - nearest["cy"]
    )

    if distance > threshold:
        return None, distance

    return nearest, distance


# ============================================================
# 星期列边界
# ============================================================

def calculate_day_boundaries(
    day_columns
):

    day_columns = sorted(
        day_columns,
        key=lambda x:
        x["cx"]
    )

    centers = [
        x["cx"]
        for x in day_columns
    ]

    if len(centers) >= 2:

        gaps = [
            centers[i]
            - centers[i - 1]
            for i in range(
                1,
                len(centers)
            )
        ]

        typical_gap = median(
            gaps
        )

    else:

        typical_gap = 180


    boundaries = []

    # 最左边
    boundaries.append(
        centers[0]
        - typical_gap / 2
    )

    # 中间
    for i in range(
        1,
        len(centers)
    ):

        boundaries.append(
            (
                centers[i - 1]
                + centers[i]
            ) / 2
        )

    # 最右边
    boundaries.append(
        centers[-1]
        + typical_gap / 2
    )

    return (
        centers,
        boundaries,
        typical_gap
    )


# ============================================================
# 根据BBox判断属于哪一天
#
# 支持合并单元格。
# ============================================================

def assign_days(
    item,
    day_columns
):
    """
    判断 OCR 文字属于哪个星期列。

    与旧版不同：
    如果文字完全位于课程网格之外，
    直接返回 []，不再强行分配给最近的一列。

    同时支持跨列的合并单元格。
    """

    centers, boundaries, cell_width = (
        calculate_day_boundaries(
            day_columns
        )
    )

    x1 = item["box"][0]
    x2 = item["box"][2]

    # --------------------------------------------------------
    # 课程网格最左、最右边界
    # --------------------------------------------------------

    grid_left = boundaries[0]
    grid_right = boundaries[-1]

    # --------------------------------------------------------
    # 关键修正：
    #
    # OCR框必须与课程网格存在实际水平重叠。
    #
    # 例如 complex1：
    #
    # “上午”
    # bbox ≈ [48, 441, 93, 517]
    #
    # 课程网格左边界 ≈ 248
    #
    # 完全没有重叠 → []
    #
    # 不再把“上午”塞到周一。
    # --------------------------------------------------------

    horizontal_overlap = max(
        0,
        min(
            x2,
            grid_right
        )
        -
        max(
            x1,
            grid_left
        )
    )

    if horizontal_overlap <= 0:
        return []

    # --------------------------------------------------------
    # 计算与每个星期列的重叠
    # --------------------------------------------------------

    overlaps = []

    for i in range(
        len(centers)
    ):

        left = boundaries[i]
        right = boundaries[i + 1]

        overlap = max(
            0,
            min(
                x2,
                right
            )
            -
            max(
                x1,
                left
            )
        )

        overlaps.append(
            overlap
        )

    bbox_width = max(
        1,
        x2 - x1
    )

    # --------------------------------------------------------
    # 合并单元格：
    #
    # 一个文字如果横跨两个或多个星期列，
    # 而且对每列都有足够重叠，
    # 则认为它属于合并单元格。
    # --------------------------------------------------------

    if bbox_width > (
        cell_width * 1.45
    ):

        selected = []

        for i, overlap in enumerate(
            overlaps
        ):

            if overlap >= (
                cell_width * 0.35
            ):

                selected.append(i)

        if selected:
            return selected

    # --------------------------------------------------------
    # 普通单元格：
    # 使用文字中心点
    # --------------------------------------------------------

    cx = item["cx"]

    for i in range(
        len(centers)
    ):

        if (
            boundaries[i]
            <= cx
            < boundaries[i + 1]
        ):

            return [i]

    # --------------------------------------------------------
    # 如果文字与网格有重叠，
    # 但是中心点刚好落在边界附近，
    # 才允许找最近列。
    #
    # 注意：这里已经经过了上面的
    # “课程网格水平重叠”检查。
    # --------------------------------------------------------

    nearest = min(
        range(len(centers)),
        key=lambda i:
        abs(
            cx
            - centers[i]
        )
    )

    return [nearest]

# ============================================================
# 合并单元格中的OCR文字
# ============================================================

def join_cell_items(items):

    if not items:
        return ""

    items = sorted(
        items,
        key=lambda x: (
            x["cy"],
            x["cx"]
        )
    )

    result_parts = []

    seen = set()

    for item in items:

        text = normalize_text(
            item["text"]
        )

        if not text:
            continue

        # 避免完全重复
        if text in seen:
            continue

        seen.add(text)

        result_parts.append(
            text
        )

    result = ""

    for text in result_parts:

        if not result:
            result = text
            continue

        # 英文前后加空格
        if (
            re.search(
                r"[A-Za-z0-9]$",
                result
            )
            and
            re.match(
                r"^[A-Za-z0-9]",
                text
            )
        ):
            result += " " + text
        else:
            result += text

    return result


# ============================================================
# 主程序
# ============================================================

print("=" * 78)
print("PaddleOCR V4 课程网格解析器")
print("=" * 78)

start_init = time.perf_counter()


ocr = PaddleOCR(
    text_detection_model_name="PP-OCRv5_mobile_det",
    text_recognition_model_name="PP-OCRv5_mobile_rec",
    use_doc_orientation_classify=False,
    use_doc_unwarping=False,
    use_textline_orientation=False,
    device="cpu",
    engine="paddle",
)


print(
    f"OCR 初始化完成："
    f"{time.perf_counter() - start_init:.2f} 秒"
)

print()


# ============================================================
# OCR
# ============================================================

print(
    f"正在识别："
    f"{IMAGE_PATH}"
)

start_ocr = time.perf_counter()

results = ocr.predict(
    IMAGE_PATH
)

ocr_time = (
    time.perf_counter()
    - start_ocr
)

print(
    f"OCR 识别完成："
    f"{ocr_time:.2f} 秒"
)

print()


if not results:
    raise RuntimeError(
        "OCR 没有返回结果"
    )


res = results[0]

data = res.json

if (
    isinstance(data, dict)
    and "res" in data
):
    data = data["res"]


texts = data.get(
    "rec_texts",
    []
)

scores = data.get(
    "rec_scores",
    []
)

boxes = data.get(
    "rec_boxes",
    [])


# ============================================================
# 整理OCR项目
# ============================================================

items = []

for i, text in enumerate(texts):

    text = normalize_text(
        text
    )

    if not text:
        continue

    box = box_to_list(
        boxes[i]
    )

    cx, cy = get_center(
        box
    )

    score = (
        float(scores[i])
        if i < len(scores)
        else 0
    )

    items.append({
        "index": i,
        "text": text,
        "score": score,
        "box": box,
        "cx": cx,
        "cy": cy
    })


print(
    f"OCR 文本区域："
    f"{len(items)} 个"
)

print()


# ============================================================
# 找星期
# ============================================================

day_columns = detect_day_columns(
    items
)


if not day_columns:

    raise RuntimeError(
        "无法确定星期列。"
        "当前图片既没有检测到明确的"
        "星期表头，也无法可靠推断列。"
    )


day_columns.sort(
    key=lambda x:
    x["cx"]
)


print("=" * 78)
print("检测到的星期列")
print("=" * 78)

for day in day_columns:

    print(
        f"{day['label']:<5}"
        f" X={day['cx']:.1f}"
        f" 来源={day['source']}"
    )

print()


# ============================================================
# 找节次
# ============================================================

periods = detect_periods(
    items,
    day_columns
)


if not periods:

    print(
        "警告：没有检测到明确的节次编号。"
    )

    print(
        "当前V4不会擅自把普通文字行当成课程节次。"
    )

    print(
        "因此本次无法可靠生成标准课程表。"
    )

    print(
        "建议下一步增加表格结构识别作为备用方案。"
    )

    sys.exit(1)


print("=" * 78)
print("检测到的有效节次")
print("=" * 78)

for period in periods:

    print(
        f"{period['label']:<6}"
        f" 原始={period['raw_label']:<5}"
        f" Y={period['cy']:.1f}"
        f" 置信度={period['score']:.3f}"
    )

print()


# ============================================================
# 节次匹配距离
# ============================================================

period_threshold = (
    calculate_period_threshold(
        periods
    )
)


print(
    f"课程文字节次匹配范围："
    f"±{period_threshold:.1f}px"
)

print()


# ============================================================
# 标记表头和节次文字
# ============================================================

day_header_indices = set(
    item["index"]
    for item in items
    if any(
        abs(
            item["cx"]
            - day["cx"]
        ) < 1
        and
        abs(
            item["cy"]
            - (
                day["cy"]
                if day["cy"] is not None
                else item["cy"]
            )
        ) < 1
        for day in day_columns
    )
)


period_indices = set()


for item in items:

    parsed = parse_period(
        item["text"]
    )

    if parsed is None:
        continue

    # 找最近一个节次锚点
    nearest = min(
        periods,
        key=lambda p:
        abs(
            item["cy"]
            - p["cy"]
        )
    )

    if abs(
        item["cy"]
        - nearest["cy"]
    ) <= 28:

        period_indices.add(
            item["index"]
        )


# ============================================================
# 提取真正靠近节次的课程文字
# ============================================================

valid_course_items = []

ignored_items = []


for item in items:

    index = item[
        "index"
    ]

    text = item[
        "text"
    ]


    # 星期表头
    if index in day_header_indices:
        continue


    # 节次文字
    if index in period_indices:
        continue


    # 任何节次编号表达式
    # 即使没有成功进入period_indices
    # 也不要把它当课程
    if parse_period(
        text
    ) is not None:
        continue


    # obvious非课程
    if is_obvious_non_course(
        text
    ):

        ignored_items.append({
            "text":
                text,
            "reason":
                "明显非课程信息"
        })

        continue


    # 将文字匹配到最近节次
    period, distance = (
        assign_period(
            item,
            periods,
            period_threshold
        )
    )


    # ------------------------------------------------------------
    # 第一关：
    # 必须属于某个真实节次
    # ------------------------------------------------------------

    if period is None:

        ignored_items.append({
            "text":
                text,

            "reason":
                "不属于任何有效节次",

            "cy":
                round(
                    item["cy"],
                    1
                ),

            "cx":
                round(
                    item["cx"],
                    1
                )
        })

        continue


    # ------------------------------------------------------------
    # 第二关：
    # 必须真正落在课程网格的横向区域
    # ------------------------------------------------------------

    day_indices = assign_days(
        item,
        day_columns
    )


    if not day_indices:

        ignored_items.append({
            "text":
                text,

            "reason":
                "位于课程网格之外",

            "cx":
                round(
                    item["cx"],
                    1
                ),

            "cy":
                round(
                    item["cy"],
                    1
                )
        })

        continue


    # ------------------------------------------------------------
    # 两个条件都满足：
    # 才是真正的课程文字
    # ------------------------------------------------------------

    valid_course_items.append({
        **item,

        "period_number":
            period["number"],

        "period_label":
            period["label"],

        "period_y":
            period["cy"],

        "period_distance":
            distance,

        # 提前保存星期列
        "day_indices":
            day_indices
    })


# ============================================================
# 打印被忽略的文字
# ============================================================

print("=" * 78)
print("被过滤/忽略的非课程文字")
print("=" * 78)

if ignored_items:

    for item in ignored_items:

        print(
            f"{item['text']:<30}"
            f" | {item['reason']}"
        )

else:

    print("没有")


print()


# ============================================================
# 创建最终课程表
# ============================================================

schedule = []


for period in periods:

    row = {
        "period_number":
            period["number"],

        "period_label":
            period["label"],

        "cells": {}
    }


    for day in day_columns:

        row["cells"][
            day["label"]
        ] = {
            "text": "",
            "min_confidence": None,
            "avg_confidence": None,
            "bbox": None,
            "ocr_items": []
        }


    schedule.append(
        row
    )


# ============================================================
# 将课程文字放入星期列
# ============================================================

for item in valid_course_items:

    day_indices = item[
        "day_indices"
    ]


    # 找对应课程行
    schedule_row = None

    for row in schedule:

        if (
            row["period_number"]
            == item["period_number"]
        ):
            schedule_row = row
            break


    if schedule_row is None:
        continue


    for day_index in day_indices:

        day = day_columns[
            day_index
        ]

        label = day[
            "label"
        ]


        cell = schedule_row[
            "cells"
        ][label]


        cell.setdefault(
            "_items",
            []
        )

        cell["_items"].append(
            item
        )


# ============================================================
# 合并每个单元格
# ============================================================

for row in schedule:

    for day in day_columns:

        label = day[
            "label"
        ]

        cell = row[
            "cells"
        ][label]


        cell_items = cell.pop(
            "_items",
            []
        )


        if not cell_items:
            continue


        text = join_cell_items(
            cell_items
        )


        scores = [
            float(
                item["score"]
            )
            for item
            in cell_items
        ]


        min_score = min(
            scores
        )

        avg_score = (
            sum(scores)
            / len(scores)
        )


        x1 = min(
            item["box"][0]
            for item
            in cell_items
        )

        y1 = min(
            item["box"][1]
            for item
            in cell_items
        )

        x2 = max(
            item["box"][2]
            for item
            in cell_items
        )

        y2 = max(
            item["box"][3]
            for item
            in cell_items
        )


        cell["text"] = text

        cell[
            "min_confidence"
        ] = round(
            min_score,
            4
        )

        cell[
            "avg_confidence"
        ] = round(
            avg_score,
            4
        )

        cell[
            "bbox"
        ] = [
            round(x1, 1),
            round(y1, 1),
            round(x2, 1),
            round(y2, 1)
        ]

        cell[
            "ocr_items"
        ] = [
            {
                "text":
                    item["text"],

                "confidence":
                    round(
                        float(
                            item["score"]
                        ),
                        4
                    ),

                "bbox": [
                    round(v, 1)
                    for v
                    in item["box"]
                ]
            }
            for item
            in cell_items
        ]


# ============================================================
# 输出标准课程表
# ============================================================

print()

print("=" * 120)
print("V4 标准课程表")
print("=" * 120)

print(
    "\t".join(
        ["节次"]
        + [
            day["label"]
            for day
            in day_columns
        ]
    )
)


for row in schedule:

    values = [
        row["period_label"]
    ]

    for day in day_columns:

        text = row[
            "cells"
        ][
            day["label"]
        ]["text"]

        values.append(
            text
        )

    print(
        "\t".join(values)
    )


# ============================================================
# 低置信度单元格
# ============================================================

low_confidence = []


for row in schedule:

    for day in day_columns:

        cell = row[
            "cells"
        ][
            day["label"]
        ]


        if not cell["text"]:
            continue


        if (
            cell["min_confidence"]
            < LOW_CONFIDENCE_THRESHOLD
        ):

            low_confidence.append({
                "period":
                    row["period_number"],

                "day":
                    day["label"],

                "text":
                    cell["text"],

                "confidence":
                    cell[
                        "min_confidence"
                    ]
            })


# ============================================================
# 创建输出目录
# ============================================================

os.makedirs(
    OUTPUT_DIR,
    exist_ok=True
)


# ============================================================
# CSV
# ============================================================

csv_path = os.path.join(
    OUTPUT_DIR,
    "course_schedule_v4.csv"
)


with open(
    csv_path,
    "w",
    newline="",
    encoding="utf-8-sig"
) as f:

    writer = csv.writer(f)

    writer.writerow(
        ["节次"]
        + [
            day["label"]
            for day
            in day_columns
        ]
    )


    for row in schedule:

        writer.writerow(
            [
                row["period_label"]
            ]
            + [
                row["cells"][
                    day["label"]
                ]["text"]
                for day
                in day_columns
            ]
        )


# ============================================================
# JSON
# ============================================================

json_path = os.path.join(
    OUTPUT_DIR,
    "course_schedule_v4.json"
)


json_data = {
    "source_image":
        IMAGE_PATH,

    "ocr_time_seconds":
        round(
            ocr_time,
            3
        ),

    "days":
        [
            {
                "label":
                    day["label"],
                "center_x":
                    round(
                        day["cx"],
                        1
                    ),
                "source":
                    day["source"]
            }
            for day
            in day_columns
        ],

    "periods":
        [
            {
                "number":
                    period["number"],

                "label":
                    period["label"],

                "raw_label":
                    period["raw_label"],

                "y":
                    round(
                        period["cy"],
                        1
                    )
            }
            for period
            in periods
        ],

    "schedule":
        schedule,

    "low_confidence":
        low_confidence,

    "ignored_items":
        ignored_items
}


with open(
    json_path,
    "w",
    encoding="utf-8"
) as f:

    json.dump(
        json_data,
        f,
        ensure_ascii=False,
        indent=2
    )


# ============================================================
# 低置信度 CSV
# ============================================================

low_path = os.path.join(
    OUTPUT_DIR,
    "low_confidence_v4.csv"
)


with open(
    low_path,
    "w",
    newline="",
    encoding="utf-8-sig"
) as f:

    writer = csv.writer(f)

    writer.writerow([
        "节次",
        "星期",
        "课程",
        "最低置信度"
    ])

    for item in low_confidence:

        writer.writerow([
            f"第{item['period']}节",
            item["day"],
            item["text"],
            item["confidence"]
        ])


# ============================================================
# 完成
# ============================================================

print()

print("=" * 78)
print("处理完成")
print("=" * 78)

print(
    f"OCR耗时："
    f"{ocr_time:.2f} 秒"
)

print(
    f"星期数量："
    f"{len(day_columns)}"
)

print(
    f"有效节次："
    f"{len(periods)}"
)

print(
    f"有效课程OCR区域："
    f"{len(valid_course_items)}"
)

print(
    f"低置信度单元格："
    f"{len(low_confidence)}"
)

print()

print(
    f"CSV：{csv_path}"
)

print(
    f"JSON：{json_path}"
)

print(
    f"低置信度：{low_path}"
)

print("=" * 78)