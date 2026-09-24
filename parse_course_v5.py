import csv
import json
import math
import os
import re
import sys
import time
import unicodedata
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from statistics import median

import numpy as np

from PIL import Image, ImageEnhance
from paddleocr import PaddleOCR


# ============================================================
# 配置
# ============================================================

# 默认直接运行时使用 test.png。
# 如果通过 BAT 拖入图片，则由 main() 用命令行参数覆盖。
IMAGE_PATH = "test.png"

# output：只放最终交付的 XLSX
OUTPUT_DIR = "output"

# debug_output：放 JSON、低置信度记录、方向矫正图片、方向候选图片等
DEBUG_OUTPUT_DIR = "debug_output"

LOW_CONFIDENCE_THRESHOLD = 0.90

# OCR文字与节次中心的最大匹配距离
PERIOD_MATCH_RATIO = 0.45

# 课程行内部文字允许的Y距离
ROW_CLUSTER_MAX_DISTANCE = 45

# ============================================================
# V5.2 图像预处理
# ============================================================

# 黑边判断：边缘像素中低亮度像素占比达到此值时，认为这一行/列可能是黑边。
BLACK_BORDER_DARK_RATIO = 0.90

# “黑色”阈值，0~255。
BLACK_BORDER_LUMINANCE = 55

# 单次最多裁掉图片尺寸的比例，防止异常图片被过度裁剪。
MAX_BORDER_CROP_RATIO = 0.35

# 小图的目标最小边长。低于此值时自动放大。
UPSCALE_TARGET_MIN_DIM = 1400

# 最大放大倍数。
UPSCALE_MAX_FACTOR = 2.0

# 放大后最长边的上限，避免极大图片造成不必要的 CPU 开销。
UPSCALE_MAX_LONG_SIDE = 3200


# ============================================================
# 非课程信息
# ============================================================

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

    "上午",
    "下午",
    "中午",
    "早读",
    "晨读",
}


NON_COURSE_KEYWORDS = [
    "备注",
    "说明",
    "课程安排说明",
    "每学期安排",
    "每周安排",
]


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
    text = text.strip()

    if text.isdigit():
        return int(text)

    if text in CHINESE_DIGITS:
        return CHINESE_DIGITS[text]

    if "十" not in text:
        return None

    if text == "十":
        return 10

    if text.startswith("十"):
        tail = text[1:]

        if tail in CHINESE_DIGITS:
            return 10 + CHINESE_DIGITS[tail]

    if text.endswith("十"):
        head = text[:-1]

        if head in CHINESE_DIGITS:
            return CHINESE_DIGITS[head] * 10

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
# 基础工具
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

    return [float(x) for x in box]


def get_center(box):
    x1, y1, x2, y2 = box

    return (
        (x1 + x2) / 2,
        (y1 + y2) / 2
    )


# ============================================================
# 节次解析
# ============================================================

def parse_period(text):
    """
    支持：

    1
    2
    第1节
    第一节
    第十二节
    3-4
    第一～二节
    """

    text = normalize_text(text)

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

    start = chinese_number_to_int(
        m.group(1)
    )

    if start is None:
        return None

    end_raw = m.group(2)

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
            if start != end
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
        r"([一二三四五六日天1-7])$",
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
# V5.2 图像预处理
# ============================================================

def _find_dark_border_crop(image):
    """
    从四个边缘自动寻找大面积连续黑边。

    使用缩小后的 NumPy 灰度矩阵一次计算行/列暗像素比例，
    避免 Python 逐像素循环影响处理速度。

    特点：
    - 只看靠近图片边缘的连续整行/整列；
    - 不会因为课程表内部存在黑色文字或黑色表格线就整体裁剪；
    - 单边最多裁掉 MAX_BORDER_CROP_RATIO。
    """

    width, height = image.size

    # 缩小后检测，降低计算量。
    max_probe = 900

    scale = min(
        1.0,
        max_probe / max(width, height)
    )

    if scale < 1.0:

        probe = image.resize(
            (
                max(1, int(width * scale)),
                max(1, int(height * scale))
            ),
            Image.Resampling.BILINEAR
        )

    else:

        probe = image

    gray_array = np.asarray(
        probe.convert("L"),
        dtype=np.uint8
    )

    dark_mask = (
        gray_array
        <= BLACK_BORDER_LUMINANCE
    )

    # 每一列 / 每一行中暗像素的比例。
    column_dark_ratio = (
        dark_mask.mean(axis=0)
    )

    row_dark_ratio = (
        dark_mask.mean(axis=1)
    )

    pw = gray_array.shape[1]
    ph = gray_array.shape[0]

    max_left = int(
        pw * MAX_BORDER_CROP_RATIO
    )

    max_top = int(
        ph * MAX_BORDER_CROP_RATIO
    )

    def find_left_run(ratios, max_count):
        count = 0

        limit = min(
            len(ratios),
            max_count
        )

        while (
            count < limit
            and
            ratios[count]
            >= BLACK_BORDER_DARK_RATIO
        ):
            count += 1

        return count

    def find_right_run(ratios, max_count):
        count = 0

        limit = min(
            len(ratios),
            max_count
        )

        while (
            count < limit
            and
            ratios[
                len(ratios) - 1 - count
            ]
            >= BLACK_BORDER_DARK_RATIO
        ):
            count += 1

        return count

    left = find_left_run(
        column_dark_ratio,
        max_left
    )

    right = find_right_run(
        column_dark_ratio,
        max_left
    )

    top = find_left_run(
        row_dark_ratio,
        max_top
    )

    bottom = find_right_run(
        row_dark_ratio,
        max_top
    )

    # 防止四边连续黑色时把整个图片裁空。
    if (
        left + right >= pw
        or
        top + bottom >= ph
    ):
        left = 0
        right = 0
        top = 0
        bottom = 0

    inv = 1.0 / scale

    crop_left = int(
        round(left * inv)
    )

    crop_top = int(
        round(top * inv)
    )

    crop_right = int(
        round((pw - right) * inv)
    )

    crop_bottom = int(
        round((ph - bottom) * inv)
    )

    crop_left = max(
        0,
        min(
            crop_left,
            width - 1
        )
    )

    crop_top = max(
        0,
        min(
            crop_top,
            height - 1
        )
    )

    crop_right = max(
        crop_left + 1,
        min(
            crop_right,
            width
        )
    )

    crop_bottom = max(
        crop_top + 1,
        min(
            crop_bottom,
            height
        )
    )

    changed = (
        crop_left > 0
        or
        crop_top > 0
        or
        crop_right < width
        or
        crop_bottom < height
    )

    if not changed:

        return image, {
            "cropped": False,
            "crop_box": [
                0,
                0,
                width,
                height
            ]
        }

    # 给真正内容留一点边缘，避免刚好贴着裁剪线。
    pad_x = max(
        2,
        int(width * 0.005)
    )

    pad_y = max(
        2,
        int(height * 0.005)
    )

    crop_left = max(
        0,
        crop_left - pad_x
    )

    crop_top = max(
        0,
        crop_top - pad_y
    )

    crop_right = min(
        width,
        crop_right + pad_x
    )

    crop_bottom = min(
        height,
        crop_bottom + pad_y
    )

    cropped = image.crop(
        (
            crop_left,
            crop_top,
            crop_right,
            crop_bottom
        )
    )

    return cropped, {
        "cropped": True,
        "crop_box": [
            crop_left,
            crop_top,
            crop_right,
            crop_bottom
        ]
    }


def _upscale_if_needed(image):
    """
    小图片自动放大，正常大图不处理。

    放大倍数：
        由最小边长决定；
        最大不超过 UPSCALE_MAX_FACTOR；
        同时限制最长边 UPSCALE_MAX_LONG_SIDE。
    """

    width, height = image.size
    min_dim = min(
        width,
        height
    )

    if min_dim >= UPSCALE_TARGET_MIN_DIM:
        return image, {
            "upscaled": False,
            "scale": 1.0
        }

    factor = (
        UPSCALE_TARGET_MIN_DIM
        / min_dim
    )

    factor = min(
        factor,
        UPSCALE_MAX_FACTOR
    )

    # 限制最长边。
    if max(width, height) * factor > UPSCALE_MAX_LONG_SIDE:
        factor = (
            UPSCALE_MAX_LONG_SIDE
            /
            max(width, height)
        )

    factor = max(
        1.0,
        factor
    )

    if factor <= 1.001:
        return image, {
            "upscaled": False,
            "scale": 1.0
        }

    new_size = (
        max(
            1,
            int(round(width * factor))
        ),
        max(
            1,
            int(round(height * factor))
        )
    )

    return (
        image.resize(
            new_size,
            Image.Resampling.LANCZOS
        ),
        {
            "upscaled": True,
            "scale": round(factor, 3)
        }
    )


def _enhance_for_ocr(image, should_enhance):
    """
    轻度增强，不做激进二值化。
    """

    if not should_enhance:
        return image

    enhanced = ImageEnhance.Contrast(
        image
    ).enhance(1.08)

    enhanced = ImageEnhance.Sharpness(
        enhanced
    ).enhance(1.15)

    return enhanced


def preprocess_input_image(input_path):
    """
    V5.2 主预处理：

        原图
          ↓
        自动去黑边
          ↓
        小图按需放大
          ↓
        小图/放大图轻度增强
          ↓
        保存到 debug_output

    只生成一个预处理版本，不额外运行第二套 OCR。
    """

    original = Image.open(
        input_path
    ).convert("RGB")

    original_size = original.size

    cropped, crop_info = (
        _find_dark_border_crop(
            original
        )
    )

    resized, resize_info = (
        _upscale_if_needed(
            cropped
        )
    )

    # 只在图像确实被裁剪或放大的情况下进行轻度增强。
    processed = _enhance_for_ocr(
        resized,
        crop_info["cropped"]
        or
        resize_info["upscaled"]
    )

    os.makedirs(
        DEBUG_OUTPUT_DIR,
        exist_ok=True
    )

    output_path = os.path.join(
        DEBUG_OUTPUT_DIR,
        "_preprocessed_input.png"
    )

    processed.save(
        output_path
    )

    return output_path, {
        "original_size": list(
            original_size
        ),
        "processed_size": list(
            processed.size
        ),
        "cropped": crop_info[
            "cropped"
        ],
        "crop_box": crop_info[
            "crop_box"
        ],
        "upscaled": resize_info[
            "upscaled"
        ],
        "scale": resize_info[
            "scale"
        ],
        "output_path": output_path
    }


# ============================================================
# 非课程判断
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
# Y 聚类
# ============================================================

def cluster_by_y(items, tolerance):

    if not items:
        return []

    items = sorted(
        items,
        key=lambda x: x["cy"]
    )

    groups = [
        [items[0]]
    ]

    for item in items[1:]:

        current = groups[-1]

        current_y = median(
            x["cy"]
            for x in current
        )

        if abs(
            item["cy"] - current_y
        ) <= tolerance:

            current.append(item)

        else:

            groups.append([
                item
            ])

    return groups


# ============================================================
# 1D KMeans
# ============================================================

def kmeans_1d(values, k):

    if len(values) < k:
        return None

    values = sorted(
        float(v)
        for v in values
    )

    centers = []

    for i in range(k):

        if k == 1:
            pos = 0
        else:
            pos = (
                i
                * (len(values) - 1)
                / (k - 1)
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

            index = min(
                range(k),
                key=lambda i:
                abs(
                    value
                    - centers[i]
                )
            )

            clusters[index].append(
                value
            )

        new_centers = []

        for i, cluster in enumerate(
            clusters
        ):

            if cluster:

                new_centers.append(
                    sum(cluster)
                    / len(cluster)
                )

            else:

                new_centers.append(
                    centers[i]
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
        key=lambda x:
        x[0]
    )

    return (
        [x[0] for x in pairs],
        [x[1] for x in pairs]
    )


# ============================================================
# 星期中文名
# ============================================================

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
# 自动检测星期列
# ============================================================

def detect_day_columns(items):

    if not items:
        return []

    max_y = max(
        item["cy"]
        for item in items
    )

    # --------------------------------------------------------
    # 方案1：星期一 / 周一 / Monday
    # --------------------------------------------------------

    explicit = []

    for item in items:

        number = parse_explicit_weekday(
            item["text"]
        )

        if number is None:
            continue

        if item["cy"] > max_y * 0.45:
            continue

        explicit.append({
            **item,
            "day_number": number
        })


    if len(explicit) >= 3:

        best_by_day = {}

        for item in explicit:

            number = item[
                "day_number"
            ]

            if (
                number not in best_by_day
                or
                item["score"]
                >
                best_by_day[number]["score"]
            ):

                best_by_day[number] = item


        if len(best_by_day) >= 3:

            result = []

            for number, item in (
                best_by_day.items()
            ):

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


            result.sort(
                key=lambda x:
                x["cx"]
            )

            return result


    # --------------------------------------------------------
    # 方案2：数字 1 2 3 4 5
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

        numeric.append(
            item
        )


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

        values = [
            int(
                x["text"]
            )
            for x in group
        ]

        values = list(
            dict.fromkeys(
                values
            )
        )

        if len(values) < 3:
            continue

        if values != list(
            range(
                1,
                len(values) + 1
            )
        ):
            continue

        candidates.append(
            group
        )


    if candidates:

        best = max(
            candidates,
            key=len
        )

        return [
            {
                "day_number":
                    int(
                        item["text"]
                    ),

                "label":
                    (
                        "周"
                        +
                        day_to_chinese(
                            int(
                                item["text"]
                            )
                        )
                    ),

                "cx":
                    item["cx"],

                "cy":
                    item["cy"],

                "source":
                    item["text"]
            }

            for item in best
        ]


    # --------------------------------------------------------
    # 方案3：中文 一 二 三 四 五
    # --------------------------------------------------------

    chinese_map = {
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

        if text not in chinese_map:
            continue

        if item["cy"] > max_y * 0.45:
            continue

        chinese_items.append(
            item
        )


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

        values = [
            chinese_map[
                x["text"]
            ]
            for x in group
        ]

        if values == list(
            range(
                1,
                len(values) + 1
            )
        ):

            candidates.append(
                group
            )


    if candidates:

        best = max(
            candidates,
            key=len
        )

        return [
            {
                "day_number":
                    chinese_map[
                        item["text"]
                    ],

                "label":
                    (
                        "周"
                        +
                        day_to_chinese(
                            chinese_map[
                                item["text"]
                            ]
                        )
                    ),

                "cx":
                    item["cx"],

                "cy":
                    item["cy"],

                "source":
                    item["text"]
            }

            for item in best
        ]


    # --------------------------------------------------------
    # 方案4：完全没有星期表头
    # --------------------------------------------------------

    candidates = []

    for item in items:

        if item["cy"] < max_y * 0.20:
            continue

        if parse_period(
            item["text"]
        ):
            continue

        if is_obvious_non_course(
            item["text"]
        ):
            continue

        candidates.append(item)


    if len(candidates) >= 15:

        values = [
            x["cx"]
            for x in candidates
        ]

        cluster_candidates = []

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
                -
                centers[i - 1]
                for i
                in range(
                    1,
                    len(centers)
                )
            ]

            if not gaps:
                continue

            average_gap = (
                sum(gaps)
                / len(gaps)
            )

            if average_gap <= 0:
                continue

            std = math.sqrt(
                sum(
                    (
                        gap
                        - average_gap
                    ) ** 2
                    for gap in gaps
                )
                / len(gaps)
            )

            cv = (
                std
                / average_gap
            )

            singleton_rate = (
                sum(
                    1
                    for cluster
                    in clusters
                    if len(cluster) <= 1
                )
                /
                len(clusters)
            )

            score = (
                cv
                + singleton_rate
                * 0.5
            )

            cluster_candidates.append(
                (
                    score,
                    centers
                )
            )


        if cluster_candidates:

            _, centers = min(
                cluster_candidates,
                key=lambda x:
                x[0]
            )

            return [
                {
                    "day_number":
                        i + 1,

                    "label":
                        (
                            "周"
                            +
                            day_to_chinese(
                                i + 1
                            )
                        ),

                    "cx":
                        center,

                    "cy":
                        None,

                    "source":
                        "自动推断"
                }

                for i, center
                in enumerate(
                    centers
                )
            ]

    return []


# ============================================================
# 自动检测节次
# ============================================================

def detect_periods(
    items,
    day_columns
):

    if not day_columns:
        return []

    day_centers = sorted(
        x["cx"]
        for x in day_columns
    )

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


    # 左侧节次区域
    grid_left = (
        day_centers[0]
        - column_gap
        * 0.72
    )


    max_y = max(
        item["cy"]
        for item in items
    )


    candidates = []

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

        candidates.append({
            **item,
            **parsed
        })


    if not candidates:
        return []


    # 同一节次附近重复OCR，只留一个
    groups = cluster_by_y(
        candidates,
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


    if not anchors:
        return []


    # 判断是不是：
    #
    # 1 2 3 4 5 6 7
    #
    # 如果是，则直接使用编号。
    #
    # 如果是：
    #
    # 1 2 3 4 1 2 3 4
    #
    # 则转换为：
    #
    # 1 2 3 4 5 6 7 8
    starts = [
        x["start"]
        for x in anchors
    ]


    use_raw_number = (
        len(set(starts))
        == len(starts)
        and
        all(
            starts[i]
            > starts[i - 1]
            for i
            in range(
                1,
                len(starts)
            )
        )
    )


    periods = []

    next_number = 1

    for anchor in anchors:

        span = (
            anchor["end"]
            -
            anchor["start"]
            + 1
        )


        if use_raw_number:

            number = anchor[
                "start"
            ]

        else:

            number = next_number


        periods.append({

            "number":
                number,

            "label":
                f"第{number}节",

            "raw_label":
                anchor["label"],

            "start":
                anchor["start"],

            "end":
                anchor["end"],

            "cx":
                anchor["cx"],

            "cy":
                anchor["cy"],

            "score":
                anchor["score"]
        })


        next_number = (
            number
            + span
        )


    return periods


# ============================================================
# 节次匹配距离
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

    gaps = [
        gap
        for gap in gaps
        if gap > 5
    ]

    if not gaps:
        return 45

    threshold = (
        median(gaps)
        * PERIOD_MATCH_RATIO
    )

    threshold = max(
        30,
        threshold
    )

    threshold = min(
        65,
        threshold
    )

    return threshold


def assign_period(
    item,
    periods,
    threshold
):

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

        return (
            None,
            distance
        )

    return (
        nearest,
        distance
    )


# ============================================================
# 星期边界
# ============================================================

def calculate_day_boundaries(
    day_columns
):

    columns = sorted(
        day_columns,
        key=lambda x:
        x["cx"]
    )

    centers = [
        x["cx"]
        for x in columns
    ]

    if len(centers) >= 2:

        gaps = [
            centers[i]
            -
            centers[i - 1]
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

    boundaries.append(
        centers[0]
        - typical_gap / 2
    )

    for i in range(
        1,
        len(centers)
    ):

        boundaries.append(
            (
                centers[i - 1]
                +
                centers[i]
            )
            / 2
        )


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
# 星期归格
# ============================================================

def assign_days(
    item,
    day_columns
):

    centers, boundaries, cell_width = (
        calculate_day_boundaries(
            day_columns
        )
    )


    x1 = item["box"][0]
    x2 = item["box"][2]

    grid_left = boundaries[0]
    grid_right = boundaries[-1]


    # --------------------------------------------------------
    # 核心：
    # 完全不在课程网格内 → 直接忽略
    # --------------------------------------------------------

    overlap = max(
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


    if overlap <= 0:
        return []


    bbox_width = max(
        1,
        x2 - x1
    )


    # --------------------------------------------------------
    # 跨列合并单元格
    # --------------------------------------------------------

    overlaps = []

    for i in range(
        len(centers)
    ):

        left = boundaries[i]
        right = boundaries[i + 1]

        value = max(
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
            value
        )


    if bbox_width > (
        cell_width * 1.45
    ):

        selected = []

        for i, value in enumerate(
            overlaps
        ):

            if value >= (
                cell_width * 0.35
            ):

                selected.append(
                    i
                )

        if selected:

            return selected


    # --------------------------------------------------------
    # 普通格子
    # --------------------------------------------------------

    cx = item["cx"]

    for i in range(
        len(centers)
    ):

        if (
            boundaries[i]
            <= cx
            <
            boundaries[i + 1]
        ):

            return [i]


    # --------------------------------------------------------
    # 只有在与网格有水平重叠的情况下，
    # 才允许最近列兜底。
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
# 教师姓名过滤
# ============================================================

# 常见中文姓氏。
# 这里不建立固定“教师名单”，只用于判断一个短文本是否“像人名”。
COMMON_SURNAMES = set(
    "赵钱孙李周吴郑王"
    "冯陈褚卫蒋沈韩杨"
    "朱秦尤许何吕施张"
    "孔曹严华金魏陶姜"
    "戚谢邹喻柏水窦章"
    "云苏潘葛奚范彭郎"
    "鲁韦昌马苗凤花方"
    "俞任袁柳酆鲍史唐"
    "费廉岑薛雷贺倪汤"
    "滕殷罗毕郝邬安常"
    "乐于时傅皮卞齐康"
    "伍余元卜顾孟平黄"
    "和穆萧尹姚邵湛汪"
    "祁毛禹狄米贝明臧"
    "计伏成戴谈宋茅庞"
    "熊纪舒屈项祝董梁"
    "杜阮蓝闵席季麻强"
    "贾路娄危江童颜郭"
    "梅盛林刁钟徐邱骆"
    "高夏蔡田樊胡凌霍"
    "虞万支柯昝管卢莫"
    "经房裘缪干解应宗"
    "丁宣贲邓郁单杭洪"
    "包诸左石崔吉钮龚"
    "程嵇邢滑裴陆荣翁"
    "荀羊於惠甄曲家封"
    "芮羿储靳汲邴糜松"
    "井段富巫乌焦巴弓"
    "牧隗山谷车侯宓蓬"
    "全郗班仰秋仲伊宫"
    "宁仇栾暴甘钭厉戎"
    "祖武符刘景詹束龙"
    "叶幸司韶郜黎蓟薄"
    "印宿白怀蒲邰从鄂"
    "索咸籍赖卓蔺屠蒙"
    "池乔阴郁胥能苍双"
    "闻莘党翟谭贡劳逄"
    "姬申扶堵冉宰郦雍"
    "郤璩桑桂濮牛寿通"
    "边扈燕冀郏浦尚农"
    "温别庄晏柴瞿阎充"
    "慕连茹习宦艾鱼容"
    "向古易慎戈廖庾终"
    "暨居衡步都耿满弘"
    "匡国文寇广禄阙东"
    "欧殳沃利蔚越夔隆"
    "师巩厍聂晁勾敖融"
    "冷訾辛阚那简饶空"
    "曾毋沙乜养鞠须丰"
    "巢关蒯相查后荆红"
    "游竺权逯盖益桓公"
)

# 明确是课程/课程组成部分的短文本。
COURSE_NAME_EXACT = {
    "语文", "数学", "英语", "外语", "科学", "体育", "音乐",
    "美术", "艺术", "书法", "劳动", "信息科技", "信息技术",
    "道德与法治", "品德与社会", "综合实践", "综合实践活动",
    "心理健康", "班队", "体活", "校本课程", "校本课", "班队会", "班队课",
    "班会", "体育与健康", "唱游", "造型", "阅读", "写字",
    "地方课程", "传统文化", "国学", "计算机", "技术",
}


def is_name_like(text):
    """
    判断一个完整 OCR 文本是否像中文姓名。
    """

    text = normalize_text(text)
    text = re.sub(r"\s+", "", text)

    if text in COURSE_NAME_EXACT:
        return False

    if not re.fullmatch(
        r"[\u4e00-\u9fff]{2,3}",
        text
    ):
        return False

    if text[0] not in COMMON_SURNAMES:
        return False

    return True


def extract_edge_name_candidates(text):
    """
    从 OCR 文本首尾寻找教师姓名候选。

    重点处理：
        语文陈雅萍
        陈雅萍语文
        综合实践活动谭明霞
        校本课(体育与健康)艾淑玮
        姜家欢唱游·乐潘姝玥
    """

    text = normalize_text(text)
    compact = re.sub(r"\s+", "", text)

    if not compact:
        return []

    candidates = []

    # “陈雅萍老师”
    match = re.search(
        r"([\u4e00-\u9fff]{2,3})老师$",
        compact
    )

    if match:
        name = match.group(1)

        if is_name_like(name):
            candidates.append(
                (name, "suffix")
            )

    # 末尾人名：课程 + 教师
    for length in (3, 2):

        if len(compact) < length:
            continue

        candidate = compact[-length:]

        if is_name_like(candidate):
            candidates.append(
                (candidate, "suffix")
            )

    # 开头人名：教师 + 课程
    for length in (3, 2):

        if len(compact) < length:
            continue

        candidate = compact[:length]

        if is_name_like(candidate):
            candidates.append(
                (candidate, "prefix")
            )

    result = []
    seen = set()

    for name, position in candidates:

        key = (
            name,
            position
        )

        if key in seen:
            continue

        seen.add(key)

        result.append(
            (name, position)
        )

    return result


def infer_teacher_names(items):
    """
    从整张课表的有效 OCR 项目中推断教师姓名。

    重要：
    2字中文文本本身存在较大歧义，例如“班队”。
    因此：
      - 3字姓名可以直接作为强候选；
      - 2字姓名若出现在“课程+姓名”末尾，可以作为候选；
      - 单独出现的2字姓名，必须结合“位于同一课程格、且明显在另一文字下方”
        或“重复出现”等证据，避免误删课程词。
    """

    exact_counts = Counter()
    suffix_counts = Counter()
    prefix_counts = Counter()
    edge_positions = {}

    # 单独出现的 2 字姓名，需要保存位置，后面结合课程格判断。
    standalone_two_char = []

    for item in items:

        text = normalize_text(
            item.get(
                "text",
                ""
            )
        )

        compact = re.sub(
            r"\s+",
            "",
            text
        )

        if not compact:
            continue

        if is_name_like(compact):

            exact_counts[compact] += 1

            if len(compact) == 2:
                standalone_two_char.append(item)

        candidates = extract_edge_name_candidates(
            compact
        )

        names_in_item = set()

        for name, position in candidates:

            names_in_item.add(name)

            edge_positions.setdefault(
                name,
                []
            ).append(position)

            if position == "suffix":
                suffix_counts[name] += 1

            elif position == "prefix":
                prefix_counts[name] += 1

        # 一个 OCR 框里同时出现两个姓名，
        # 例如：姜家欢唱游·乐潘姝玥。
        if len(names_in_item) >= 2:

            for name in names_in_item:

                edge_positions.setdefault(
                    name,
                    []
                ).append(
                    "multi_name"
                )

    teacher_names = set()

    # --------------------------------------------------------
    # 1. 3字独立姓名：强候选
    # 2字独立文本不能仅凭“像姓名”直接删除。
    # --------------------------------------------------------
    for name in exact_counts:

        if len(name) >= 3:
            teacher_names.add(name)

    # --------------------------------------------------------
    # 2. 单独出现的2字姓名：
    # 如果它在同一个课程格里明显位于另一段文字下方，
    # 则很像“课程在上、教师在下”的版式。
    # --------------------------------------------------------
    for item in standalone_two_char:

        name = re.sub(
            r"\s+",
            "",
            item["text"]
        )

        same_cell_items = []

        for other in items:

            if other is item:
                continue

            if (
                other.get("period_number")
                !=
                item.get("period_number")
            ):
                continue

            item_days = set(
                item.get(
                    "day_indices",
                    []
                )
            )

            other_days = set(
                other.get(
                    "day_indices",
                    []
                )
            )

            if not (
                item_days
                &
                other_days
            ):
                continue

            same_cell_items.append(
                other
            )

        has_upper_text = any(
            other["cy"]
            <
            item["cy"] - 8
            for other
            in same_cell_items
            if other.get("text")
        )

        if has_upper_text:
            teacher_names.add(name)

    # --------------------------------------------------------
    # 3. “课程 + 教师”：
    # 尾部姓名是目前最可靠的组合形式。
    # --------------------------------------------------------
    for name in suffix_counts:

        # 两字姓名也允许，例如：
        # 英语赵静
        teacher_names.add(name)

    # --------------------------------------------------------
    # 4. “教师 + 课程”：
    # 需要更多证据，避免课程词误判。
    # --------------------------------------------------------
    for name, count in prefix_counts.items():

        positions = edge_positions.get(
            name,
            []
        )

        if (
            len(name) >= 3
            and
            (
                exact_counts.get(name, 0) > 0
                or
                count >= 2
                or
                "multi_name" in positions
            )
        ):
            teacher_names.add(name)

    return teacher_names

def remove_teacher_names(
    text,
    teacher_names
):
    """
    从课程文本中删除教师姓名。
    """

    text = normalize_text(text)

    text = re.sub(
        r"\s+",
        "",
        text
    )

    if not text or not teacher_names:
        return text

    # 长姓名优先。
    for name in sorted(
        teacher_names,
        key=len,
        reverse=True
    ):

        if not name:
            continue

        text = text.replace(
            name,
            ""
        )

    # 清理尾部“老师 / 教师”标记。
    text = re.sub(
        r"(?:教师|老师)[:：]?$",
        "",
        text
    )

    return normalize_text(text).strip()


# ============================================================
# 单元格文字合并
# ============================================================

def join_cell_items(
    items,
    teacher_names=None
):
    """
    合并课程格中的 OCR 项目，同时删除教师姓名。
    """

    if not items:
        return ""

    teacher_names = (
        teacher_names
        if teacher_names is not None
        else set()
    )

    items = sorted(
        items,
        key=lambda x: (
            x["cy"],
            x["cx"]
        )
    )

    parts = []
    seen = set()

    # 第一层：逐 OCR 框清洗
    for item in items:

        text = normalize_text(
            item["text"]
        )

        if not text:
            continue

        # OCR 单独识别出一整行教师姓名。
        if (
            text in teacher_names
            and
            is_name_like(text)
        ):
            continue

        # OCR 把课程和教师粘成一个框。
        text = remove_teacher_names(
            text,
            teacher_names
        )

        if not text:
            continue

        if text in seen:
            continue

        seen.add(text)
        parts.append(text)

    # 第二层：合并剩余课程文字
    result = ""

    for text in parts:

        if not result:
            result = text
            continue

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

    return result.strip()


# ============================================================
# 方向模型结果解析
# ============================================================

# ============================================================
# 方向模型结果解析
# ============================================================





# ============================================================
# OCR
# ============================================================

def run_ocr(
    ocr,
    image_path
):

    results = ocr.predict(
        image_path
    )

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
        []
    )


    items = []

    for i, text in enumerate(
        texts
    ):

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
            float(
                scores[i]
            )
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


    return (
        items,
        res
    )


# ============================================================
# 解析一个方向
# ============================================================

def parse_orientation(
    items
):

    if not items:

        return {
            "success": False,
            "reason": "没有OCR结果"
        }


    # --------------------------------------------------------
    # 星期
    # --------------------------------------------------------

    day_columns = (
        detect_day_columns(
            items
        )
    )


    if len(day_columns) < 3:

        return {
            "success": False,
            "reason":
                f"星期列不足：{len(day_columns)}"
        }


    # --------------------------------------------------------
    # 节次
    # --------------------------------------------------------

    periods = detect_periods(
        items,
        day_columns
    )


    if len(periods) < 1:

        return {
            "success": False,
            "reason":
                "没有检测到节次"
        }


    # --------------------------------------------------------
    # 匹配距离
    # --------------------------------------------------------

    threshold = (
        calculate_period_threshold(
            periods
        )
    )


    # --------------------------------------------------------
    # 课程文字
    # --------------------------------------------------------

    valid_items = []

    ignored = []


    for item in items:

        text = item[
            "text"
        ]


        # 星期表头
        if parse_explicit_weekday(
            text
        ) is not None:

            continue


        # 纯数字星期
        if re.fullmatch(
            r"[1-7]",
            text
        ):

            continue


        # 节次
        if parse_period(
            text
        ) is not None:

            continue


        # 明确非课程
        if is_obvious_non_course(
            text
        ):

            ignored.append({
                "text":
                    text,
                "reason":
                    "非课程关键词"
            })

            continue


        # Y匹配节次
        period, distance = (
            assign_period(
                item,
                periods,
                threshold
            )
        )


        if period is None:

            ignored.append({
                "text":
                    text,
                "reason":
                    "不属于课程节次",
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


        # X匹配课程网格
        day_indices = assign_days(
            item,
            day_columns
        )


        if not day_indices:

            ignored.append({
                "text":
                    text,
                "reason":
                    "位于课程网格外",
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


        valid_items.append({
            **item,

            "period_number":
                period["number"],

            "period_label":
                period["label"],

            "period_y":
                period["cy"],

            "period_distance":
                distance,

            "day_indices":
                day_indices
        })


    # --------------------------------------------------------
    # 推断教师姓名
    # --------------------------------------------------------
    #
    # 只使用已经位于课程网格内的有效 OCR 项目。
    # 网格外的“上午 / 下午”等文字不会进入教师候选。
    #
    teacher_names = infer_teacher_names(
        valid_items
    )

    # --------------------------------------------------------
    # 建立课程表
    # --------------------------------------------------------

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

            row[
                "cells"
            ][
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


    # --------------------------------------------------------
    # 放入课程
    # --------------------------------------------------------

    for item in valid_items:

        row = None

        for schedule_row in schedule:

            if (
                schedule_row[
                    "period_number"
                ]
                ==
                item[
                    "period_number"
                ]
            ):

                row = schedule_row

                break


        if row is None:
            continue


        for day_index in (
            item["day_indices"]
        ):

            day_label = (
                day_columns[
                    day_index
                ]["label"]
            )


            cell = row[
                "cells"
            ][
                day_label
            ]


            cell.setdefault(
                "_items",
                []
            )


            cell[
                "_items"
            ].append(
                item
            )


    # --------------------------------------------------------
    # 合并单元格
    # --------------------------------------------------------

    for row in schedule:

        for day in day_columns:

            label = day[
                "label"
            ]

            cell = row[
                "cells"
            ][
                label
            ]


            cell_items = cell.pop(
                "_items",
                []
            )


            if not cell_items:
                continue


            text = join_cell_items(
                cell_items,
                teacher_names
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
                /
                len(scores)
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


            cell[
                "text"
            ] = text

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
                        round(
                            value,
                            1
                        )
                        for value
                        in item["box"]
                    ]
                }

                for item
                in cell_items
            ]


    # --------------------------------------------------------
    # 有效课程数量
    # --------------------------------------------------------

    non_empty = 0

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


            non_empty += 1


            if (
                cell[
                    "min_confidence"
                ]
                <
                LOW_CONFIDENCE_THRESHOLD
            ):

                low_confidence.append({
                    "period":
                        row[
                            "period_number"
                        ],

                    "day":
                        day["label"],

                    "text":
                        cell["text"],

                    "confidence":
                        cell[
                            "min_confidence"
                        ]
                })


    # --------------------------------------------------------
    # 结构评分
    #
    # 用于判断课程表结构是否合理。
    # --------------------------------------------------------

    expected = (
        len(periods)
        * len(day_columns)
    )


    fill_rate = (
        non_empty
        / expected
        if expected > 0
        else 0
    )


    # 有星期、有节次、有一定课程数量
    structure_score = (
        min(
            len(day_columns),
            7
        )
        * 10
        +
        min(
            len(periods),
            12
        )
        * 5
        +
        min(
            non_empty,
            50
        )
        * 0.5
        +
        fill_rate
        * 10
    )


    # --------------------------------------------------------
    # 最低条件：
    #
    # 星期 >= 3
    # 节次 >= 1
    # 至少有一个课程格
    # --------------------------------------------------------

    success = (
        len(day_columns) >= 3
        and
        len(periods) >= 1
        and
        non_empty >= 1
    )


    return {
        "success":
            success,

        "reason":
            (
                "OK"
                if success
                else "课程网格为空"
            ),

        "day_columns":
            day_columns,

        "periods":
            periods,

        "schedule":
            schedule,

        "ignored":
            ignored,

        "valid_item_count":
            len(valid_items),

        "teacher_names":
            sorted(
                teacher_names
            ),

        "non_empty_cells":
            non_empty,

        "low_confidence":
            low_confidence,

        "structure_score":
            structure_score
    }




# ============================================================
# 清理最终输出目录
# ============================================================

def prepare_output_dir():
    """
    output 是最终交付目录。
    每次运行前清理其中的旧文件，确保最终只留下 XLSX。
    """

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    for name in os.listdir(
        OUTPUT_DIR
    ):

        path = os.path.join(
            OUTPUT_DIR,
            name
        )

        if os.path.isfile(path):
            try:
                os.remove(path)
            except OSError:
                pass


def save_xlsx(
    schedule,
    day_columns,
    xlsx_path
):
    """
    将标准课程表保存为 XLSX。
    """

    wb = Workbook()

    ws = wb.active
    ws.title = "课程表"

    headers = (
        ["节次"]
        +
        [
            day["label"]
            for day
            in day_columns
        ]
    )

    ws.append(headers)

    for row in schedule:

        ws.append(
            [
                row["period_label"]
            ]
            +
            [
                row["cells"][
                    day["label"]
                ]["text"]
                for day
                in day_columns
            ]
        )

    # 表头与单元格采用适合人工校对的基础格式。
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row in ws.iter_rows(
        min_row=2
    ):
        for cell in row:
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

    # 第一列节次稍窄，其余课程列适当加宽。
    ws.column_dimensions["A"].width = 10

    for column in range(
        2,
        len(headers) + 1
    ):
        column_letter = ws.cell(
            row=1,
            column=column
        ).column_letter

        ws.column_dimensions[
            column_letter
        ].width = 24

    ws.freeze_panes = "B2"
    ws.sheet_view.showGridLines = True

    # 根据文字量设置一个适中的行高。
    for row_index in range(
        2,
        ws.max_row + 1
    ):
        max_length = max(
            len(str(ws.cell(
                row=row_index,
                column=column
            ).value or ""))
            for column in range(
                1,
                ws.max_column + 1
            )
        )

        ws.row_dimensions[
            row_index
        ].height = max(
            24,
            min(
                60,
                18 + max_length * 0.8
            )
        )

    wb.save(
        xlsx_path
    )


# ============================================================
# 保存结果
# ============================================================

def save_outputs(
    selected,
    input_path,
    orientation_score,
    preprocess_info=None
):

    parsed = selected[
        "parsed"
    ]

    day_columns = parsed[
        "day_columns"
    ]

    periods = parsed[
        "periods"
    ]

    schedule = parsed[
        "schedule"
    ]

    low_confidence = parsed[
        "low_confidence"
    ]


    prepare_output_dir()

    os.makedirs(
        DEBUG_OUTPUT_DIR,
        exist_ok=True
    )


    # --------------------------------------------------------
    # XLSX：最终交付文件，只放在 output
    # --------------------------------------------------------

    xlsx_path = os.path.join(
        OUTPUT_DIR,
        "course_schedule.xlsx"
    )

    save_xlsx(
        schedule,
        day_columns,
        xlsx_path
    )


    # --------------------------------------------------------
    # JSON：调试/分析文件
    # --------------------------------------------------------

    json_path = os.path.join(
        DEBUG_OUTPUT_DIR,
        "course_schedule_v5.json"
    )


    json_data = {

        "source_image":
            input_path,

        "input_orientation":
            "upright",

        "ocr_time_seconds":
            round(
                selected[
                    "ocr_time"
                ],
                3
            ),

        "preprocess":
            (
                {
                    **preprocess_info
                }
                if preprocess_info is not None
                else {}
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

        "filtered_teacher_names":
            parsed.get(
                "teacher_names",
                []
            ),


        "ignored_items":
            parsed[
                "ignored"
            ]
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


    # --------------------------------------------------------
    # 低置信度 CSV：调试/分析文件
    # --------------------------------------------------------

    low_path = os.path.join(
        DEBUG_OUTPUT_DIR,
        "low_confidence_v5.csv"
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


    # --------------------------------------------------------
    # 最终预处理图片：调试文件
    # --------------------------------------------------------

    final_image_path = os.path.join(
        DEBUG_OUTPUT_DIR,
        "preprocessed_final.png"
    )


    Image.open(
        selected["image_path"]
    ).save(
        final_image_path
    )


    return {
        "xlsx":
            xlsx_path,

        "json":
            json_path,

        "low_confidence":
            low_path,

        "image":
            final_image_path
    }


# 主程序
# ============================================================

def main():

    global IMAGE_PATH

    # --------------------------------------------------------
    # 命令行图片参数
    # --------------------------------------------------------
    #
    # 直接运行：
    #     python parse_course_v5.py
    #     -> 使用 test.png
    #
    # 拖拽图片到 BAT：
    #     BAT 会把图片完整路径作为第一个参数传进来。
    #
    # 支持路径中包含空格。
    #
    if len(sys.argv) > 1:

        IMAGE_PATH = os.path.abspath(
            sys.argv[1]
        )

    else:

        IMAGE_PATH = os.path.abspath(
            IMAGE_PATH
        )


    print("=" * 78)
    print("PaddleOCR V5 课程表解析器")
    print("=" * 78)


    print(
        f"输入图片：{IMAGE_PATH}"
    )


    # --------------------------------------------------------
    # 检查图片
    # --------------------------------------------------------

    if not os.path.exists(
        IMAGE_PATH
    ):

        print(
            f"找不到图片："
            f"{IMAGE_PATH}"
        )

        sys.exit(1)


    # output 只作为最终交付目录。
    # 运行开始时先清空旧文件，避免解析失败时留下上一张课表的旧 XLSX。
    prepare_output_dir()

    os.makedirs(
        DEBUG_OUTPUT_DIR,
        exist_ok=True
    )


    # --------------------------------------------------------
    # V5.2 图像预处理
    # --------------------------------------------------------

    print()
    print(
        "正在预处理图片..."
    )

    preprocess_start = (
        time.perf_counter()
    )

    preprocessed_path, preprocess_info = (
        preprocess_input_image(
            IMAGE_PATH
        )
    )

    preprocess_time = (
        time.perf_counter()
        - preprocess_start
    )

    print(
        f"预处理耗时："
        f"{preprocess_time:.2f} 秒"
    )

    print(
        f"原图尺寸："
        f"{preprocess_info['original_size'][0]}"
        f"x"
        f"{preprocess_info['original_size'][1]}"
    )

    print(
        f"处理后尺寸："
        f"{preprocess_info['processed_size'][0]}"
        f"x"
        f"{preprocess_info['processed_size'][1]}"
    )

    if preprocess_info["cropped"]:
        print(
            f"检测到黑边并自动裁剪："
            f"{preprocess_info['crop_box']}"
        )
    else:
        print(
            "未检测到需要裁剪的黑边"
        )

    if preprocess_info["upscaled"]:
        print(
            f"图片已放大："
            f"{preprocess_info['scale']:.3f}x"
        )
    else:
        print(
            "无需放大图片"
        )


    # --------------------------------------------------------
    # 加载 OCR 模型
    # --------------------------------------------------------

    print()
    print(
        "正在加载 OCR 模型..."
    )

    ocr_start = (
        time.perf_counter()
    )

    ocr = PaddleOCR(

        text_detection_model_name=
            "PP-OCRv5_mobile_det",

        text_recognition_model_name=
            "PP-OCRv5_mobile_rec",

        use_doc_orientation_classify=
            False,

        use_doc_unwarping=
            False,

        use_textline_orientation=
            False,

        device=
            "cpu",

        engine=
            "paddle"
    )

    print(
        "OCR模型加载完成："
        f"{time.perf_counter() - ocr_start:.2f}"
        " 秒"
    )


    # --------------------------------------------------------
    # 单次 OCR + 课程表解析
    # --------------------------------------------------------

    start_total = (
        time.perf_counter()
    )

    ocr_run_start = (
        time.perf_counter()
    )

    items, res = run_ocr(
        ocr,
        preprocessed_path
    )

    ocr_time = (
        time.perf_counter()
        - ocr_run_start
    )

    parsed = parse_orientation(
        items
    )

    print()
    print("=" * 78)
    print("OCR与课程表解析")
    print("=" * 78)

    print(
        f"OCR耗时："
        f"{ocr_time:.2f} 秒"
    )

    print(
        f"文字区域："
        f"{len(items)}"
    )

    print(
        f"星期列："
        f"{len(parsed.get('day_columns', []))}"
    )

    print(
        f"节次："
        f"{len(parsed.get('periods', []))}"
    )

    print(
        f"有效课程格："
        f"{parsed.get('non_empty_cells', 0)}"
    )

    print(
        f"结构分数："
        f"{parsed.get('structure_score', 0):.2f}"
    )

    total_time = (
        time.perf_counter()
        - start_total
    )

    selected = {
        "angle": 0,
        "items": items,
        "result": res,
        "parsed": parsed,
        "ocr_time": ocr_time,
        "image_path": preprocessed_path
    }

    if not parsed.get("success"):
        print()
        print("=" * 78)
        print("无法可靠解析课程表")
        print("=" * 78)

        print(
            f"原因："
            f"{parsed.get('reason', '未知错误')}"
        )

        print(
            "程序不会输出新的有效课程表，"
            "请检查图片是否清晰、课程表是否为正向。"
        )

        sys.exit(2)

    # --------------------------------------------------------
    # 输出
    # --------------------------------------------------------

    outputs = save_outputs(
        selected,
        IMAGE_PATH,
        0.0,
        preprocess_info
    )


    # --------------------------------------------------------
    # 打印最终课程表
    # --------------------------------------------------------

    parsed = selected[
        "parsed"
    ]

    day_columns = parsed[
        "day_columns"
    ]

    schedule = parsed[
        "schedule"
    ]


    print()
    print("=" * 120)
    print("V5 标准课程表")
    print("=" * 120)


    print(
        "\t".join(
            ["节次"]
            +
            [
                day["label"]
                for day
                in day_columns
            ]
        )
    )


    for row in schedule:

        print(
            "\t".join(
                [
                    row[
                        "period_label"
                    ]
                ]
                +
                [
                    row[
                        "cells"
                    ][
                        day["label"]
                    ][
                        "text"
                    ]
                    for day
                    in day_columns
                ]
            )
        )


    # --------------------------------------------------------
    # 统计
    # --------------------------------------------------------

    non_empty_cells = parsed[
        "non_empty_cells"
    ]

    low_count = len(
        parsed[
            "low_confidence"
        ]
    )


    print()
    print("=" * 78)
    print("处理完成")
    print("=" * 78)


    print(
        "图片方向：默认正向，不进行自动旋转"
    )


    print(
        f"OCR耗时："
        f"{selected['ocr_time']:.2f} 秒"
    )


    print(
        f"总处理时间："
        f"{total_time:.2f} 秒"
    )


    print(
        f"星期数量："
        f"{len(day_columns)}"
    )


    print(
        f"有效节次："
        f"{len(schedule)}"
    )


    print(
        f"非空课程格："
        f"{non_empty_cells}"
    )


    print(
        f"低置信度课程格："
        f"{low_count}"
    )


    print()
    print(
        f"XLSX："
        f"{outputs['xlsx']}"
    )

    print(
        f"调试JSON："
        f"{outputs['json']}"
    )

    print(
        f"低置信度："
        f"{outputs['low_confidence']}"
    )

    print(
        f"正向图片："
        f"{outputs['image']}"
    )

    print("=" * 78)


if __name__ == "__main__":
    main()