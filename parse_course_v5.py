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
from openpyxl.styles import Font, Alignment, PatternFill
from statistics import median

import numpy as np

try:
    import cv2
except ImportError:
    cv2 = None

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

# ------------------------------------------------------------
# 局部单元格补识别
# ------------------------------------------------------------

# 每张课表最多做多少个局部单元格补识别，避免速度失控。
LOCAL_OCR_MAX_CELLS = 12

# 空单元格只有在所在行已有足够课程内容时，才进入候选。
LOCAL_OCR_EMPTY_ROW_MIN_FILLED = 2

# 空单元格接受局部 OCR 的最低平均置信度。
LOCAL_OCR_EMPTY_MIN_CONFIDENCE = 0.78

# 非空但可疑单元格接受局部 OCR 的最低平均置信度。
LOCAL_OCR_REPAIR_MIN_CONFIDENCE = 0.72

# 文本短于等于该长度时，认为有较高概率存在截断。
LOCAL_OCR_SHORT_TEXT_LENGTH = 2

# 局部裁剪向内缩，避免表格边线进入 OCR。
LOCAL_OCR_INSET_RATIO_X = 0.035
LOCAL_OCR_INSET_RATIO_Y = 0.08

# OCR文字与节次中心的最大匹配距离
PERIOD_MATCH_RATIO = 0.45

# 课程行内部文字允许的Y距离
ROW_CLUSTER_MAX_DISTANCE = 45

# ============================================================
# V5.3 图像预处理
# ============================================================

# 黑边判断：边缘像素中低亮度像素占比达到此值时，认为这一行/列可能是黑边。
BLACK_BORDER_DARK_RATIO = 0.75

# “黑色”阈值，0~255。
BLACK_BORDER_LUMINANCE = 55

# 如果整行/整列的平均亮度足够低，也视为黑边。
BLACK_BORDER_MEAN_LUMINANCE = 90

# 单边最多裁掉图片尺寸的比例。
MAX_BORDER_CROP_RATIO = 0.45

# 自动裁边后，至少保留原尺寸的这一比例。
MIN_REMAINING_RATIO = 0.20

# 白边/空白边检测的最小亮度对比。
WHITE_BORDER_MIN_CONTRAST = 12

# 行/列中至少有这么多比例的前景像素，才认为这一行/列属于内容。
CONTENT_DARK_RATIO = 0.003

# 小图的目标最小边长。低于此值时自动放大。
UPSCALE_TARGET_MIN_DIM = 1600

# 最大放大倍数。
UPSCALE_MAX_FACTOR = 2.5

# 放大后最长边的上限，避免极大图片造成不必要的 CPU 开销。
UPSCALE_MAX_LONG_SIDE = 3600

# ------------------------------------------------------------
# OCR 图像增强
#
# 参考成熟深度学习 OCR 的做法：
# - 默认保留原始 RGB；
# - 只在检测到“整体对比度明显偏低”时做轻度增强；
# - 增强只作用于亮度通道，不破坏原始颜色；
# - 不默认执行中值滤波、局部背景模糊、强制灰度、激进二值化。
#
# 百分位对比度 = P90 - P10。
# 该值越低，通常表示整张图片的文字/背景亮度分离越弱。
# ------------------------------------------------------------

# 用于快速估计对比度的最大探测边长。
OCR_CONTRAST_PROBE_MAX = 900

# P90-P10 低于此值时，才认为需要轻度增强。
# 数值越高，越容易触发增强。
OCR_LOW_CONTRAST_THRESHOLD = 65

# 低对比度图片的轻度亮度对比增强倍率。
# 建议先在 1.10~1.25 范围内调整。
OCR_LOW_CONTRAST_FACTOR = 1.18

# ============================================================
# OCR 专用清洗
# ============================================================

# 自适应二值化窗口必须为奇数。
OCR_ADAPTIVE_BLOCK_SIZE = 31

# 自适应二值化常数。
OCR_ADAPTIVE_C = 9

# 网格线提取的最小长度（像素）。
# 课程表网格线远长于单个汉字笔画。
OCR_MIN_LINE_LENGTH = 30

# OCR 调试输入文件名。
OCR_INPUT_DEBUG_NAME = "ocr_input.png"


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
    "晨诵",
    "晨练",
}

# OCR 可能把活动名称拆成多个短框。
NON_COURSE_FRAGMENT_COMBINATIONS = {
    "午休",
    "午睡",
    "晨读",
    "晨诵",
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

    # 课程表常见写法：
    # 上午1 / 上午一 / 上1
    # 下午1 / 下午一 / 下1
    # 保留原始 1~4 编号，后续由 detect_periods() 按整张表的垂直顺序重新编号。
    session_match = re.match(
        r"^(上午|下午|上|下)\s*"
        r"([0-9]{1,2}|[一二三四五六七八九十]+)$",
        text
    )

    if session_match:
        start = chinese_number_to_int(
            session_match.group(2)
        )

        if (
            start is not None
            and
            1 <= start <= 12
        ):
            return {
                "start": start,
                "end": start,
                "label": text,
                "session": session_match.group(1)
            }

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
        ),
        "session": None
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

    # 再计算整行/整列平均亮度。
    # 对“整体很黑但混有少量亮色像素”的 JPEG / 截图黑边更稳健。
    column_mean_luminance = (
        gray_array.mean(axis=0)
    )

    row_mean_luminance = (
        gray_array.mean(axis=1)
    )

    pw = gray_array.shape[1]
    ph = gray_array.shape[0]

    max_left = int(
        pw * MAX_BORDER_CROP_RATIO
    )

    max_top = int(
        ph * MAX_BORDER_CROP_RATIO
    )

    def is_dark_edge(
        dark_ratio,
        mean_luminance
    ):
        return (
            dark_ratio
            >= BLACK_BORDER_DARK_RATIO
            or
            mean_luminance
            <= BLACK_BORDER_MEAN_LUMINANCE
        )

    def find_left_run(
        ratios,
        mean_luminances,
        max_count
    ):
        count = 0

        limit = min(
            len(ratios),
            max_count
        )

        while (
            count < limit
            and
            is_dark_edge(
                ratios[count],
                mean_luminances[count]
            )
        ):
            count += 1

        return count

    def find_right_run(
        ratios,
        mean_luminances,
        max_count
    ):
        count = 0

        limit = min(
            len(ratios),
            max_count
        )

        while (
            count < limit
            and
            is_dark_edge(
                ratios[
                    len(ratios) - 1 - count
                ],
                mean_luminances[
                    len(mean_luminances) - 1 - count
                ]
            )
        ):
            count += 1

        return count

    left = find_left_run(
        column_dark_ratio,
        column_mean_luminance,
        max_left
    )

    right = find_right_run(
        column_dark_ratio,
        column_mean_luminance,
        max_left
    )

    top = find_left_run(
        row_dark_ratio,
        row_mean_luminance,
        max_top
    )

    bottom = find_right_run(
        row_dark_ratio,
        row_mean_luminance,
        max_top
    )

    # 如果上下/左右合计裁掉后只剩极薄区域，则放弃该方向裁剪。
    if (
        pw - left - right
        < pw * MIN_REMAINING_RATIO
    ):
        left = 0
        right = 0

    if (
        ph - top - bottom
        < ph * MIN_REMAINING_RATIO
    ):
        top = 0
        bottom = 0

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
            ],
            "detected_border_runs": [
                0,
                0,
                0,
                0
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
        ],
        "detected_border_runs": [
            crop_left,
            crop_top,
            width - crop_right,
            height - crop_bottom
        ]
    }


def _find_content_crop(image):
    """
    自动裁掉四周大面积白边/浅色空白。

    思路：
    - 黑边先由 _find_dark_border_crop() 处理；
    - 这里再从剩余图片的四周背景估计“纸张/空白”的亮度；
    - 找出相对背景明显更暗的文字、表格线、印章等内容；
    - 依据行/列前景密度得到内容包围盒。

    即使原图没有黑边、只有白边，也可以自动缩紧到真正内容区域。
    """

    width, height = image.size

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

    gray = np.asarray(
        probe.convert("L"),
        dtype=np.uint8
    )

    ph, pw = gray.shape

    # 从四个角落估计空白背景亮度。
    band_x = max(1, int(pw * 0.06))
    band_y = max(1, int(ph * 0.06))

    corner_pixels = np.concatenate([
        gray[:band_y, :band_x].reshape(-1),
        gray[:band_y, pw - band_x:].reshape(-1),
        gray[ph - band_y:, :band_x].reshape(-1),
        gray[ph - band_y:, pw - band_x:].reshape(-1),
    ])

    background_luminance = float(
        np.median(corner_pixels)
    )

    threshold = (
        background_luminance
        - WHITE_BORDER_MIN_CONTRAST
    )

    threshold = max(
        120.0,
        min(
            245.0,
            threshold
        )
    )

    foreground = (
        gray < threshold
    )

    row_density = (
        foreground.mean(axis=1)
    )

    column_density = (
        foreground.mean(axis=0)
    )

    row_indices = np.where(
        row_density >= CONTENT_DARK_RATIO
    )[0]

    column_indices = np.where(
        column_density >= CONTENT_DARK_RATIO
    )[0]

    if (
        len(row_indices) == 0
        or
        len(column_indices) == 0
    ):
        return image, {
            "cropped": False,
            "crop_box": [
                0,
                0,
                width,
                height
            ],
            "background_luminance":
                round(
                    background_luminance,
                    1
                ),
            "threshold":
                round(
                    threshold,
                    1
                )
        }

    left = int(
        round(
            column_indices[0]
            / scale
        )
    )

    top = int(
        round(
            row_indices[0]
            / scale
        )
    )

    right = int(
        round(
            (column_indices[-1] + 1)
            / scale
        )
    )

    bottom = int(
        round(
            (row_indices[-1] + 1)
            / scale
        )
    )

    left = max(
        0,
        min(
            left,
            width - 1
        )
    )

    top = max(
        0,
        min(
            top,
            height - 1
        )
    )

    right = max(
        left + 1,
        min(
            right,
            width
        )
    )

    bottom = max(
        top + 1,
        min(
            bottom,
            height
        )
    )

    changed = (
        left > 0
        or
        top > 0
        or
        right < width
        or
        bottom < height
    )

    if not changed:
        return image, {
            "cropped": False,
            "crop_box": [
                0,
                0,
                width,
                height
            ],
            "background_luminance":
                round(
                    background_luminance,
                    1
                ),
            "threshold":
                round(
                    threshold,
                    1
                )
        }

    # 给边缘内容留少量安全边距。
    pad_x = max(
        3,
        int(width * 0.01)
    )

    pad_y = max(
        3,
        int(height * 0.01)
    )

    left = max(
        0,
        left - pad_x
    )

    top = max(
        0,
        top - pad_y
    )

    right = min(
        width,
        right + pad_x
    )

    bottom = min(
        height,
        bottom + pad_y
    )

    cropped = image.crop(
        (
            left,
            top,
            right,
            bottom
        )
    )

    return cropped, {
        "cropped": True,
        "crop_box": [
            left,
            top,
            right,
            bottom
        ],
        "background_luminance":
            round(
                background_luminance,
                1
            ),
        "threshold":
            round(
                threshold,
                1
            )
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
            Image.Resampling.BICUBIC
        ),
        {
            "upscaled": True,
            "scale": round(factor, 3)
        }
    )


def _measure_luminance_contrast(image):
    """
    快速测量图片整体亮度对比度。

    不对整张高分辨率图片计算百分位，而是先缩小到最多
    OCR_CONTRAST_PROBE_MAX 的长边，降低 CPU 开销。

    返回：
        {
            "p10": ...,
            "p90": ...,
            "contrast": ...
        }
    """

    width, height = image.size

    scale = min(
        1.0,
        OCR_CONTRAST_PROBE_MAX
        / max(width, height)
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

    gray = np.asarray(
        probe.convert("L"),
        dtype=np.uint8
    )

    p10, p90 = np.percentile(
        gray,
        [10, 90]
    )

    return {
        "p10": round(float(p10), 1),
        "p90": round(float(p90), 1),
        "contrast": round(
            float(p90 - p10),
            1
        )
    }


def _enhance_for_ocr(image):
    """
    V5.3 OCR 图像增强。

    核心原则：
    1. 默认保留原始 RGB；
    2. 默认不做灰度化、MedianFilter、Gaussian 背景校正或强制
       autocontrast；
    3. 先快速测量整体亮度对比度；
    4. 只有明显低对比度时，才对 Y（亮度）通道做轻度 Contrast；
    5. Cb/Cr（颜色）完全来自原图，从而保留原始颜色。

    这样更接近现代深度学习 OCR 常见的“Resize + Normalize”
    路线，同时保留 EasyOCR 一类项目的“低对比度时再增强”的思路。
    """

    contrast_info = _measure_luminance_contrast(
        image
    )

    contrast = contrast_info["contrast"]

    if contrast >= OCR_LOW_CONTRAST_THRESHOLD:
        return (
            image,
            {
                **contrast_info,
                "enhanced": False,
                "factor": 1.0,
                "reason": "对比度足够，无需额外增强"
            }
        )

    # 只处理亮度通道，保留原始色彩。
    ycbcr = image.convert(
        "YCbCr"
    )

    y, cb, cr = ycbcr.split()

    enhanced_y = ImageEnhance.Contrast(
        y
    ).enhance(
        OCR_LOW_CONTRAST_FACTOR
    )

    enhanced = Image.merge(
        "YCbCr",
        (
            enhanced_y,
            cb,
            cr
        )
    ).convert(
        "RGB"
    )

    after = _measure_luminance_contrast(
        enhanced
    )

    return (
        enhanced,
        {
            **after,
            "before_p10": contrast_info["p10"],
            "before_p90": contrast_info["p90"],
            "before_contrast": contrast_info["contrast"],
            "enhanced": True,
            "factor": OCR_LOW_CONTRAST_FACTOR,
            "reason": (
                f"低对比度，亮度通道增强 "
                f"{OCR_LOW_CONTRAST_FACTOR:.2f}x"
            )
        }
    )


def _build_ocr_input(image):
    """
    制作真正送给 PaddleOCR 的清洗版图片。

    重点处理本项目这类课程表：
    - 小字周围有压缩噪点/纸张纹理；
    - 横竖网格线很多；
    - 原图肉眼可看，但通用 OCR 检测器容易把网格和杂点当成干扰。

    只生成一个 OCR 输入，不增加第二次 OCR。
    """

    # 没有 OpenCV 时，使用轻量 PIL 方案作为后备。
    if cv2 is None:
        gray = image.convert("L")

        gray = ImageEnhance.Contrast(
            gray
        ).enhance(1.12)

        return gray.convert("RGB"), {
            "method":
                "PIL灰度+轻度对比增强",
            "opencv":
                False,
            "line_removal":
                False,
            "adaptive_threshold":
                False
        }

    rgb = np.asarray(
        image,
        dtype=np.uint8
    )

    gray = cv2.cvtColor(
        rgb,
        cv2.COLOR_RGB2GRAY
    )

    # 先去掉最明显的孤立杂点。
    denoised = cv2.medianBlur(
        gray,
        3
    )

    # 局部背景归一化：
    # 消除纸张发灰、拍照光照不均带来的大尺度亮度变化。
    background = cv2.GaussianBlur(
        denoised,
        (0, 0),
        sigmaX=7
    )

    normalized = cv2.divide(
        denoised,
        background,
        scale=255
    )

    # 先得到前景掩膜，用于提取长横线/竖线。
    binary = cv2.adaptiveThreshold(
        normalized,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        OCR_ADAPTIVE_BLOCK_SIZE,
        OCR_ADAPTIVE_C
    )

    foreground = cv2.bitwise_not(
        binary
    )

    height, width = foreground.shape

    horizontal_length = max(
        OCR_MIN_LINE_LENGTH,
        int(width * 0.04)
    )

    vertical_length = max(
        OCR_MIN_LINE_LENGTH,
        int(height * 0.04)
    )

    horizontal_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (
            horizontal_length,
            1
        )
    )

    vertical_kernel = cv2.getStructuringElement(
        cv2.MORPH_RECT,
        (
            1,
            vertical_length
        )
    )

    horizontal_lines = cv2.morphologyEx(
        foreground,
        cv2.MORPH_OPEN,
        horizontal_kernel
    )

    vertical_lines = cv2.morphologyEx(
        foreground,
        cv2.MORPH_OPEN,
        vertical_kernel
    )

    line_mask = cv2.bitwise_or(
        horizontal_lines,
        vertical_lines
    )

    # 轻微膨胀，覆盖灰色网格线边缘。
    line_mask = cv2.dilate(
        line_mask,
        np.ones(
            (2, 2),
            dtype=np.uint8
        ),
        iterations=1
    )

    # 在灰度图上修补网格线，而不是简单涂成纯白。
    cleaned = cv2.inpaint(
        normalized,
        line_mask,
        2,
        cv2.INPAINT_TELEA
    )

    # 去线后重新二值化。
    cleaned_binary = cv2.adaptiveThreshold(
        cleaned,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        OCR_ADAPTIVE_BLOCK_SIZE,
        OCR_ADAPTIVE_C
    )

    return (
        Image.fromarray(
            cleaned_binary
        ).convert("RGB"),
        {
            "method":
                "OpenCV中值去噪+局部背景归一化+表格线去除+自适应二值化",
            "opencv":
                True,
            "line_removal":
                True,
            "adaptive_threshold":
                True,
            "horizontal_kernel":
                horizontal_length,
            "vertical_kernel":
                vertical_length
        }
    )


def preprocess_input_image(input_path):
    """
    V5.3 主预处理：

        原图 RGB
          ↓
        自动去黑边
          ↓
        自动去白边/空白边
          ↓
        小图按需放大
          ↓
        检测整体对比度
          ↓
        仅低对比度时对亮度通道做轻度增强
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

    content_cropped, content_info = (
        _find_content_crop(
            cropped
        )
    )

    resized, resize_info = (
        _upscale_if_needed(
            content_cropped
        )
    )

    # OCR 专用增强：
    # 默认保持 RGB；只有低对比度图片才对亮度通道做轻度增强。
    processed, enhance_info = _enhance_for_ocr(
        resized
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

    # 额外生成一张“清洗版”供对比观察。
    # 注意：该版本可能过度处理小字，因此暂不直接作为 PaddleOCR 输入。
    ocr_image, ocr_info = _build_ocr_input(
        processed
    )

    ocr_input_path = os.path.join(
        DEBUG_OUTPUT_DIR,
        OCR_INPUT_DEBUG_NAME
    )

    ocr_image.save(
        ocr_input_path
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
        "detected_border_runs": crop_info[
            "detected_border_runs"
        ],
        "content_cropped": content_info[
            "cropped"
        ],
        "content_crop_box": content_info[
            "crop_box"
        ],
        "content_background_luminance":
            content_info[
                "background_luminance"
            ],
        "content_threshold":
            content_info[
                "threshold"
            ],
        "ocr_denoise":
            False,
        "ocr_contrast_enhanced":
            enhance_info["enhanced"],
        "ocr_contrast_reason":
            enhance_info["reason"],
        "ocr_contrast_before":
            enhance_info.get(
                "before_contrast",
                enhance_info["contrast"]
            ),
        "ocr_contrast_after":
            enhance_info["contrast"],
        "ocr_contrast_threshold":
            OCR_LOW_CONTRAST_THRESHOLD,
        "ocr_contrast_factor":
            enhance_info["factor"],
        "ocr_input_path":
            ocr_input_path,
        "ocr_actual_input_path":
            output_path,
        "ocr_cleaning_method":
            ocr_info["method"],
        "ocr_opencv":
            ocr_info["opencv"],
        "ocr_line_removal":
            ocr_info["line_removal"],
        "ocr_adaptive_threshold":
            ocr_info["adaptive_threshold"],
        "ocr_horizontal_kernel":
            ocr_info.get(
                "horizontal_kernel"
            ),
        "ocr_vertical_kernel":
            ocr_info.get(
                "vertical_kernel"
            ),
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

    compact = re.sub(
        r"\s+",
        "",
        text
    )

    if compact in NON_COURSE_EXACT:
        return True

    for keyword in NON_COURSE_KEYWORDS:
        if keyword in compact:
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
# 课程名规范化
# ============================================================

COURSE_CANONICAL_ALIASES = {
    "体育与健": "体育与健康",
    "体有与健康": "体育与健康",
    "体育与健一": "体育与健康",
    "体育与健二": "体育与健康",
    "体有": "体育",
    "道与法": "道德与法治",
    "道法": "道德与法治",
    "综合实践": "综合实践活动",
}


def canonicalize_course_text(text):
    """
    对 OCR 产生的常见截断/错字做保守修正。

    重点：
    - 先拆出课程主体与括号备注；
    - 对主体应用别名修正；
    - 再恢复括号内容；
    - 最后才做保守的前缀/子序列匹配。
    """

    text = normalize_text(text)

    text = re.sub(
        r"^[，,；;。]+|[，,；;。]+$",
        "",
        text
    )

    compact = re.sub(
        r"\s+",
        "",
        text
    )

    if not compact:
        return ""

    suffix = ""
    base = compact

    match = re.match(
        r"^(.+?)([（(].*[）)])$",
        compact
    )

    if match:
        base = match.group(1)
        suffix = match.group(2)

    # 先处理已知 OCR 错别字/截断。
    if base in COURSE_CANONICAL_ALIASES:
        return (
            COURSE_CANONICAL_ALIASES[
                base
            ]
            +
            suffix
        )

    if base in COURSE_NAME_EXACT:
        return base + suffix

    # 只对长度 >= 3 的中文课程主体进行保守模糊匹配。
    if (
        len(base) < 3
        or
        not re.fullmatch(
            r"[\u4e00-\u9fff·]+",
            base
        )
    ):
        return text

    best = None

    for course in COURSE_NAME_EXACT:

        if len(course) < 3:
            continue

        # OCR 主体是课程名前缀。
        if course.startswith(base):
            ratio = (
                len(base)
                /
                len(course)
            )

            if ratio >= 0.55:
                candidate = (
                    ratio,
                    course
                )

                if (
                    best is None
                    or
                    candidate[0] > best[0]
                ):
                    best = candidate

        # OCR 主体按顺序保留了课程名的主要汉字。
        pos = 0

        for char in course:
            if (
                pos < len(base)
                and
                char == base[pos]
            ):
                pos += 1

        if pos == len(base):
            ratio = (
                len(base)
                /
                len(course)
            )

            if ratio >= 0.60:
                candidate = (
                    ratio,
                    course
                )

                if (
                    best is None
                    or
                    candidate[0] > best[0]
                ):
                    best = candidate

    if best is not None:
        return best[1] + suffix

    return text




# ============================================================
# 从课程网格本身推断行
# ============================================================


def _detect_schedule_row_bands(
    items,
    day_columns,
    image_path
):
    """
    从真实横向表格线识别课程行，并判断是否为“合并活动行”。

    merged=True：
        行内没有正常的星期列竖线，典型例子是“午休”。

    standard row：
        有星期列竖线，即使课程文字全部漏掉，也应保留这一行。
    """

    if (
        cv2 is None
        or not image_path
        or not os.path.exists(image_path)
        or len(day_columns) < 2
    ):
        return []

    try:
        image = cv2.imread(
            image_path,
            cv2.IMREAD_GRAYSCALE
        )

        if image is None:
            return []

        height, width = image.shape

        centers = sorted(
            day["cx"]
            for day in day_columns
        )

        gaps = [
            centers[i] - centers[i - 1]
            for i in range(
                1,
                len(centers)
            )
        ]

        if not gaps:
            return []

        typical_gap = median(gaps)

        grid_left = max(
            0,
            int(
                centers[0]
                - typical_gap * 0.75
            )
        )

        grid_right = min(
            width,
            int(
                centers[-1]
                + typical_gap * 0.75
            )
        )

        roi = image[
            :,
            grid_left:grid_right
        ]

        binary = cv2.adaptiveThreshold(
            roi,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV,
            31,
            7
        )

        horizontal_length = max(
            40,
            int(
                roi.shape[1] * 0.025
            )
        )

        horizontal_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (
                horizontal_length,
                1
            )
        )

        horizontal = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            horizontal_kernel
        )

        coverage = (
            horizontal.astype(
                np.float32
            ).mean(axis=1)
        )

        candidate_rows = np.where(
            coverage >= 0.35
        )[0]

        if len(candidate_rows) == 0:
            return []

        groups = []

        for y in candidate_rows:
            y = int(y)

            if not groups:
                groups.append([y])
                continue

            if y - groups[-1][-1] <= 3:
                groups[-1].append(y)
            else:
                groups.append([y])

        lines = [
            float(median(group))
            for group in groups
        ]

        header_y = median(
            day["cy"]
            for day in day_columns
            if day.get("cy") is not None
        )

        lines = [
            line
            for line in lines
            if line > header_y + 12
        ]

        if len(lines) < 2:
            return []

        # 表格横线的相邻距离。
        line_gaps = [
            lines[i] - lines[i - 1]
            for i in range(
                1,
                len(lines)
            )
            if lines[i] - lines[i - 1] > 10
        ]

        if not line_gaps:
            return []

        typical_row_height = median(
            line_gaps
        )

        # 计算真实竖线覆盖率，用于识别合并行。
        vertical_kernel_height = max(
            12,
            int(
                typical_row_height * 0.55
            )
        )

        vertical_kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT,
            (
                1,
                vertical_kernel_height
            )
        )

        vertical = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            vertical_kernel
        )

        expected_boundaries = [
            int(
                round(
                    centers[i]
                    -
                    grid_left
                    -
                    typical_gap / 2
                )
            )
            for i in range(
                1,
                len(centers)
            )
        ]

        bands = []

        for i in range(
            len(lines) - 1
        ):

            top = lines[i]
            bottom = lines[i + 1]
            row_height = bottom - top

            if row_height < 15:
                continue

            if row_height > (
                typical_row_height * 2.6
            ):
                continue

            # 在当前 band 中检查 4 条星期列分隔线是否存在。
            internal_vertical_count = 0

            for boundary_x in expected_boundaries:
                y1 = max(
                    0,
                    int(round(top + 3))
                )

                y2 = min(
                    height,
                    int(round(bottom - 3))
                )

                if y2 <= y1:
                    continue

                x1 = max(
                    0,
                    boundary_x - 2
                )

                x2 = min(
                    roi.shape[1],
                    boundary_x + 3
                )

                if x2 <= x1:
                    continue

                score = float(
                    vertical[
                        y1:y2,
                        x1:x2
                    ].mean()
                    /
                    255.0
                )

                if score >= 0.12:
                    internal_vertical_count += 1

            bands.append({
                "top": top,
                "bottom": bottom,
                "cy": (top + bottom) / 2,
                "internal_vertical_count":
                    internal_vertical_count,
                "merged":
                    internal_vertical_count < 2
            })

        return bands

    except Exception:
        return []


def infer_periods_from_grid(
    items,
    day_columns,
    image_path=None
):
    """
    通过真实表格横线确定课程行，不依赖左侧节次标签。

    行分类：
    - 明确的晨诵/午休/课间活动行：直接过滤；
    - 前后完全没有课程文字的伪行：过滤；
    - 中间完全空白的行：保留，供局部 OCR 补识别；
    - 任意非活动课程名称默认保留，不依赖课程词库。
    """

    if not day_columns:
        return []

    columns, boundaries, cell_width = (
        calculate_day_boundaries(
            day_columns
        )
    )

    header_y = median(
        day["cy"]
        for day in day_columns
        if day.get("cy") is not None
    )

    grid_left = boundaries[0]
    grid_right = boundaries[-1]

    bands = _detect_schedule_row_bands(
        items,
        day_columns,
        image_path
    )

    if bands:

        records = []

        for band in bands:

            band_items = [
                item
                for item in items
                if (
                    band["top"] - 4
                    <= item["cy"]
                    <= band["bottom"] + 4
                    and
                    item["cx"] >= grid_left
                    and
                    item["cx"] < grid_right
                )
            ]

            ordered = sorted(
                band_items,
                key=lambda x: (
                    x["cy"],
                    x["cx"]
                )
            )

            compact_text = "".join(
                re.sub(
                    r"\s+",
                    "",
                    normalize_text(
                        item.get(
                            "text",
                            ""
                        )
                    )
                )
                for item
                in ordered
            )

            activity_hit = (
                is_obvious_non_course(
                    compact_text
                )
                or
                any(
                    fragment in compact_text
                    for fragment
                    in NON_COURSE_FRAGMENT_COMBINATIONS
                )
            )

            # 这里不依赖 is_name_like 来判断“是否课程”。
            # 课程名称可能只有2~4字，也可能恰好以常见姓氏开头。
            content_items = [
                item
                for item
                in band_items
                if (
                    normalize_text(
                        item.get(
                            "text",
                            ""
                        )
                    )
                    and
                    not (
                        re.sub(
                            r"\s+",
                            "",
                            normalize_text(
                                item.get(
                                    "text",
                                    ""
                                )
                            )
                        )
                        in {
                            "午",
                            "休",
                            "晨",
                            "读"
                        }
                    )
                    and
                    not is_obvious_non_course(
                        item.get(
                            "text",
                            ""
                        )
                    )
                )
            ]

            records.append({
                "band":
                    band,

                "band_items":
                    band_items,

                "content_items":
                    content_items,

                "activity":
                    activity_hit
            })

        # 明确活动行直接删除。
        records = [
            record
            for record in records
            if not record["activity"]
        ]

        # 前后空白伪行删除；中间空白行保留给局部 OCR。
        meaningful = [
            index
            for index, record
            in enumerate(records)
            if record["content_items"]
        ]

        if meaningful:

            records = records[
                meaningful[0]:
                meaningful[-1] + 1
            ]
        else:
            records = []

        periods = []

        for record in records:

            band = record["band"]
            band_items = record["band_items"]
            content_items = record["content_items"]

            number = len(periods) + 1

            periods.append({
                "number":
                    number,

                "label":
                    f"第{number}节",

                "raw_label":
                    None,

                "start":
                    number,

                "end":
                    number,

                "cx":
                    (
                        median(
                            item["cx"]
                            for item
                            in content_items
                        )
                        if content_items
                        else
                        median(columns)
                    ),

                "cy":
                    band["cy"],

                "score":
                    (
                        min(
                            float(
                                item["score"]
                            )
                            for item
                            in content_items
                        )
                        if content_items
                        else
                        1.0
                    ),

                "source":
                    "表格横线行推断",

                "item_count":
                    len(band_items),

                "course_item_count":
                    len(content_items),

                "row_top":
                    band["top"],

                "row_bottom":
                    band["bottom"],

                "merged":
                    band.get(
                        "merged",
                        False
                    ),

                "internal_vertical_count":
                    band.get(
                        "internal_vertical_count",
                        0
                    ),

                "empty_before_local_ocr":
                    not bool(
                        content_items
                    )
            })

        if len(periods) >= 2:
            return periods

    # 无可靠表格线时的备用方案。
    candidates = []

    for item in items:

        text = normalize_text(
            item.get(
                "text",
                ""
            )
        )

        compact_text = re.sub(
            r"\s+",
            "",
            text
        )

        if not compact_text:
            continue

        if item["cy"] <= header_y + 20:
            continue

        if parse_explicit_weekday(
            text
        ) is not None:
            continue

        if re.fullmatch(
            r"[1-7]",
            text
        ):
            continue

        if item["cx"] < grid_left:
            continue

        if item["cx"] >= grid_right:
            continue

        if parse_period(
            text
        ) is not None:
            continue

        if is_obvious_non_course(
            compact_text
        ):
            continue

        if compact_text in {
            "午",
            "休",
            "晨",
            "读"
        }:
            continue

        if not assign_days(
            item,
            day_columns
        ):
            continue

        candidates.append(item)

    if not candidates:
        return []

    heights = [
        max(
            1,
            item["box"][3] - item["box"][1]
        )
        for item in candidates
    ]

    row_tolerance = max(
        16,
        min(
            ROW_CLUSTER_MAX_DISTANCE,
            median(heights) * 1.15
        )
    )

    groups = cluster_by_y(
        candidates,
        tolerance=row_tolerance
    )

    return [
        {
            "number":
                index,

            "label":
                f"第{index}节",

            "raw_label":
                None,

            "start":
                index,

            "end":
                index,

            "cx":
                median(
                    item["cx"]
                    for item
                    in group
                ),

            "cy":
                median(
                    item["cy"]
                    for item
                    in group
                ),

            "score":
                min(
                    float(item["score"])
                    for item
                    in group
                ),

            "source":
                "课程网格OCR行推断",

            "item_count":
                len(group),

            "row_tolerance":
                row_tolerance
        }
        for index, group
        in enumerate(
            groups,
            start=1
        )
    ]

# ============================================================
# 自动检测节次
# ============================================================

def detect_periods(
    items,
    day_columns,
    image_path=None
):
    """
    节次检测。

    主方案：
        直接从课程网格中的 OCR 文字按 Y 聚类得到课程行。

    备用方案：
        如果课程网格文字不足，再尝试识别左侧节次标签。
    """

    # --------------------------------------------------------
    # 主方案：从课程网格本身推断行
    # --------------------------------------------------------

    periods = infer_periods_from_grid(
        items,
        day_columns,
        image_path
    )

    if periods:
        return periods

    # --------------------------------------------------------
    # 备用方案：左侧节次标签
    # --------------------------------------------------------

    day_centers = sorted(
        x["cx"]
        for x in day_columns
    )

    if len(day_centers) >= 2:

        gaps = [
            day_centers[i]
            -
            day_centers[i - 1]
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

        if item["cx"] >= grid_left:
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
            ==
            starts[i - 1] + 1
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

            "session":
                anchor.get(
                    "session"
                ),

            "start":
                anchor["start"],

            "end":
                anchor["end"],

            "cx":
                anchor["cx"],

            "cy":
                anchor["cy"],

            "score":
                anchor["score"],

            "source":
                "左侧节次标签",

            "item_count":
                1
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

    row_bounded = any(
        "row_top" in period
        and
        "row_bottom" in period
        for period
        in periods
    )

    for period in periods:

        if (
            "row_top" in period
            and
            "row_bottom" in period
            and
            period["row_top"] - 3
            <= item["cy"]
            <= period["row_bottom"] + 3
        ):
            return (
                period,
                0.0
            )

    if row_bounded:
        return (
            None,
            float("inf")
        )

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
    "语文", "数学", "英语", "外语", "日语", "俄语", "科学", "自然",
    "物理", "化学", "生物", "生物学", "生命科学", "地理", "历史", "历史与社会",
    "体育", "体育与健康", "体活", "体育活动", "健康教育", "音乐", "美术", "艺术",
    "舞蹈", "戏剧", "影视", "唱游", "造型", "手工", "美工", "音乐欣赏", "美术欣赏",
    "书法", "写字", "硬笔书法", "软笔书法", "阅读", "习作", "作文", "阅读与写作",
    "劳动", "劳动技术", "劳技", "劳动与技术", "通用技术", "信息技术", "信息科技",
    "计算机", "微机", "电脑", "技术",
    "道德与法治", "道法", "道与法", "品德与生活", "品德与社会", "思想品德",
    "思品", "思想政治", "思政",
    "综合实践", "综合实践活动", "研究性学习", "社会实践", "研学实践",
    "心理健康", "心理健康教育", "心理辅导", "心理",
    "安全教育", "国防教育", "环境教育", "生态文明教育", "生涯规划",
    "班队", "班队会", "班队课", "班会", "班与队", "队会", "少先队活动",
    "少先队活动课", "主题班会", "德育", "德育课",
    "校本课程", "校本课", "地方课程", "传统文化", "国学", "地方文化", "乡土课程",
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
    基于单元格版面证据推断教师姓名。

    设计原则：
    - 任何未知文本默认视为课程；
    - 姓氏表只提供弱证据；
    - 单独姓名必须同时具备“下置”以及重复/字号等第二证据；
    - 粘连姓名优先依赖尾部位置；
    - 已知课程词不会因为名字特征被删除。
    """

    if not items:
        return set()

    cell_groups = {}

    for item in items:
        for day_index in item.get(
            "day_indices",
            []
        ):
            key = (
                item.get(
                    "period_number"
                ),
                day_index
            )

            cell_groups.setdefault(
                key,
                []
            ).append(item)

    evidence = {}
    occurrences = {}
    layout_support = {}
    suffix_support = {}

    def add_evidence(
        name,
        score
    ):
        name = re.sub(
            r"\s+",
            "",
            str(name)
        )

        if (
            len(name) < 2
            or
            len(name) > 3
            or
            not re.fullmatch(
                r"[\u4e00-\u9fff]{2,3}",
                name
            )
            or
            name in COURSE_NAME_EXACT
        ):
            return

        evidence[name] = (
            evidence.get(
                name,
                0
            )
            + score
        )

    for key, group in cell_groups.items():

        ordered = sorted(
            group,
            key=lambda x: (
                x["cy"],
                x["cx"]
            )
        )

        for item in ordered:

            compact = re.sub(
                r"\s+",
                "",
                normalize_text(
                    item.get(
                        "text",
                        ""
                    )
                )
            )

            if (
                len(compact) not in (2, 3)
                or
                not re.fullmatch(
                    r"[\u4e00-\u9fff]{2,3}",
                    compact
                )
                or
                compact in COURSE_NAME_EXACT
            ):
                continue

            item_height = max(
                1,
                item["box"][3]
                -
                item["box"][1]
            )

            above = [
                other
                for other in ordered
                if other is not item
                and
                other["cy"]
                <
                item["cy"]
                -
                max(
                    6,
                    item_height * 0.35
                )
            ]

            if not above:
                continue

            min_y = min(
                other["cy"]
                for other
                in ordered
            )

            max_y = max(
                other["cy"]
                for other
                in ordered
            )

            lower_position = (
                item["cy"]
                >=
                min_y
                +
                (max_y - min_y) * 0.55
            )

            upper_heights = [
                max(
                    1,
                    other["box"][3]
                    -
                    other["box"][1]
                )
                for other
                in above
            ]

            smaller_than_course = any(
                item_height
                <=
                height * 0.90
                for height
                in upper_heights
            )

            score = 3

            if compact[0] in COMMON_SURNAMES:
                score += 1

            if lower_position:
                score += 1

            if smaller_than_course:
                score += 2
                layout_support.setdefault(
                    compact,
                    False
                )
                layout_support[
                    compact
                ] = True

            add_evidence(
                compact,
                score
            )

            occurrences.setdefault(
                compact,
                set()
            ).add(key)

    # 跨多个独立课程格重复出现，是强证据。
    for name, cells in occurrences.items():
        if len(cells) >= 2:
            add_evidence(
                name,
                min(
                    3,
                    len(cells)
                )
            )

    # 粘连文本：
    #   语文张老师 / 语文张润凝
    #   张润凝语文（前缀仅作为弱证据）
    for item in items:

        compact = re.sub(
            r"\s+",
            "",
            normalize_text(
                item.get(
                    "text",
                    ""
                )
            )
        )

        if not compact:
            continue

        for name, position in extract_edge_name_candidates(
            compact
        ):

            if name in COURSE_NAME_EXACT:
                continue

            if position == "suffix":
                add_evidence(
                    name,
                    2
                )

                suffix_support[
                    name
                ] = True

                if name[0] in COMMON_SURNAMES:
                    add_evidence(
                        name,
                        1
                    )

            elif position == "prefix":
                # 前缀形式可能与真实课程词混淆，只保留弱证据；
                # 最终必须还有重复或其它版面证据。
                add_evidence(
                    name,
                    1
                )

    teacher_names = set()

    for name, score in evidence.items():

        if score < 4:
            continue

        repeated = (
            len(
                occurrences.get(
                    name,
                    set()
                )
            )
            >= 2
        )

        # 必须至少具备：
        # 1) 多个课程格重复；
        # 2) 单元格字号/位置等额外版面证据；
        # 3) 粘连文本尾部姓名证据。
        if (
            repeated
            or
            layout_support.get(
                name,
                False
            )
            or
            suffix_support.get(
                name,
                False
            )
        ):
            teacher_names.add(
                name
            )

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

        # 独立的单字通常是 OCR 噪声/活动残片。
        compact_item = re.sub(
            r"\s+",
            "",
            text
        )

        if (
            len(compact_item) <= 1
            and
            compact_item in {
                "午",
                "休",
                "晨",
                "读",
                "语"
            }
        ):
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

        text = canonicalize_course_text(
            text
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

def _get_period_crop_bounds(
    periods,
    index,
    image_height
):
    """
    获取一个标准课程行的垂直范围。
    """

    period = periods[index]

    if (
        "row_top" in period
        and
        "row_bottom" in period
    ):
        return (
            max(
                0,
                float(
                    period["row_top"]
                )
            ),
            min(
                float(
                    period["row_bottom"]
                ),
                float(image_height)
            )
        )

    center = float(
        period["cy"]
    )

    if index > 0:
        top = (
            float(
                periods[index - 1]["cy"]
            )
            + center
        ) / 2
    else:
        gap = (
            float(
                periods[index + 1]["cy"]
            )
            -
            center
            if index + 1 < len(periods)
            else 80
        )
        top = center - gap / 2

    if index + 1 < len(periods):
        bottom = (
            center
            +
            float(
                periods[index + 1]["cy"]
            )
        ) / 2
    else:
        gap = (
            center
            -
            float(
                periods[index - 1]["cy"]
            )
            if index > 0
            else 80
        )
        bottom = center + gap / 2

    return (
        max(
            0,
            top
        ),
        min(
            float(image_height),
            bottom
        )
    )


def _get_cell_key(
    item,
    day_index
):
    return (
        item.get(
            "period_number"
        ),
        day_index
    )


def _cell_raw_text(
    items
):
    """
    在教师过滤前，只用于判断单元格是否可疑。
    """

    if not items:
        return ""

    parts = []

    for item in sorted(
        items,
        key=lambda x: (
            x["cy"],
            x["cx"]
        )
    ):
        text = re.sub(
            r"\s+",
            "",
            normalize_text(
                item.get(
                    "text",
                    ""
                )
            )
        )

        if not text:
            continue

        if is_obvious_non_course(
            text
        ):
            continue

        if text in {
            "午",
            "休",
            "晨",
            "读"
        }:
            continue

        parts.append(text)

    return "".join(parts)


def _should_refine_cell(
    text,
    confidence,
    row_filled
):
    """
    只把真正可疑的单元格送入局部 OCR。
    """

    if not text:
        return (
            row_filled
            >=
            LOCAL_OCR_EMPTY_ROW_MIN_FILLED
        )

    if len(text) <= LOCAL_OCR_SHORT_TEXT_LENGTH:
        return True

    if confidence < LOW_CONFIDENCE_THRESHOLD:
        return True

    return False


def refine_suspicious_cells(
    valid_items,
    periods,
    day_columns,
    image_path,
    ocr
):
    """
    对可疑课程格进行局部补识别。

    关键设计：
    - 仍然只初始化一次 PaddleOCR；
    - 只对少量高优先级单元格调用 predict；
    - 局部裁剪向内缩，主动避开网格线；
    - 局部结果如果更完整，则替换该单元格原始 OCR 项，
      不与旧文字重复拼接。
    """

    empty_stats = {
        "candidates": 0,
        "attempted": 0,
        "accepted": 0,
        "time_seconds": 0.0
    }

    if (
        ocr is None
        or
        not image_path
        or
        not os.path.exists(image_path)
        or
        not periods
        or
        not day_columns
    ):
        return (
            valid_items,
            empty_stats
        )

    try:
        image = Image.open(
            image_path
        ).convert("RGB")
    except Exception:
        return (
            valid_items,
            empty_stats
        )

    _, boundaries, _ = calculate_day_boundaries(
        day_columns
    )

    cell_groups = {}

    for item in valid_items:
        for day_index in item.get(
            "day_indices",
            []
        ):
            cell_groups.setdefault(
                (
                    item["period_number"],
                    day_index
                ),
                []
            ).append(item)

    candidates = []

    for period in periods:

        row_filled = 0

        for day_index in range(
            len(day_columns)
        ):
            key = (
                period["number"],
                day_index
            )

            if _cell_raw_text(
                cell_groups.get(
                    key,
                    []
                )
            ):
                row_filled += 1

        for day_index in range(
            len(day_columns)
        ):

            key = (
                period["number"],
                day_index
            )

            items = cell_groups.get(
                key,
                []
            )

            text = _cell_raw_text(
                items
            )

            scores = [
                float(
                    item.get(
                        "score",
                        0
                    )
                )
                for item
                in items
                if item.get(
                    "text"
                )
                and
                not is_obvious_non_course(
                    item.get(
                        "text",
                        ""
                    )
                )
            ]

            confidence = (
                min(scores)
                if scores
                else 0.0
            )

            if not _should_refine_cell(
                text,
                confidence,
                row_filled
            ):
                continue

            priority = 0

            if not text:
                priority += 5
                if row_filled >= 3:
                    priority += 2

            elif len(text) <= 2:
                priority += 4

            if confidence < LOW_CONFIDENCE_THRESHOLD:
                priority += 2

            if text in {
                "语",
                "读",
                "品",
                "体",
                "学",
                "市"
            }:
                priority += 3

            candidates.append({
                "period":
                    period["number"],
                "period_index":
                    periods.index(
                        period
                    ),
                "day_index":
                    day_index,
                "current_text":
                    text,
                "current_score":
                    confidence,
                "priority":
                    priority
            })

    candidates.sort(
        key=lambda item: (
            -item["priority"],
            item["current_score"],
            item["period"],
            item["day_index"]
        )
    )

    selected = candidates[
        :LOCAL_OCR_MAX_CELLS
    ]

    if not selected:
        return (
            valid_items,
            empty_stats
        )

    os.makedirs(
        os.path.join(
            DEBUG_OUTPUT_DIR,
            "local_ocr"
        ),
        exist_ok=True
    )

    refined_items = list(
        valid_items
    )

    accepted = 0
    started = time.perf_counter()

    for attempt_index, candidate in enumerate(
        selected,
        start=1
    ):

        period_index = candidate[
            "period_index"
        ]

        day_index = candidate[
            "day_index"
        ]

        row_top, row_bottom = (
            _get_period_crop_bounds(
                periods,
                period_index,
                image.height
            )
        )

        left = boundaries[
            day_index
        ]

        right = boundaries[
            day_index + 1
        ]

        cell_width = (
            right - left
        )

        cell_height = (
            row_bottom - row_top
        )

        inset_x = max(
            6,
            int(
                cell_width
                * LOCAL_OCR_INSET_RATIO_X
            )
        )

        inset_y = max(
            6,
            int(
                cell_height
                * LOCAL_OCR_INSET_RATIO_Y
            )
        )

        crop_left = int(
            max(
                0,
                left + inset_x
            )
        )

        crop_top = int(
            max(
                0,
                row_top + inset_y
            )
        )

        crop_right = int(
            min(
                image.width,
                right - inset_x
            )
        )

        crop_bottom = int(
            min(
                image.height,
                row_bottom - inset_y
            )
        )

        if (
            crop_right <= crop_left + 12
            or
            crop_bottom <= crop_top + 12
        ):
            continue

        crop = image.crop(
            (
                crop_left,
                crop_top,
                crop_right,
                crop_bottom
            )
        )

        crop_width, crop_height = crop.size

        scale = (
            520
            /
            min(
                crop_width,
                crop_height
            )
        )

        scale = max(
            1.0,
            min(
                2.5,
                scale
            )
        )

        if scale > 1.01:
            crop = crop.resize(
                (
                    int(
                        crop_width
                        * scale
                    ),
                    int(
                        crop_height
                        * scale
                    )
                ),
                Image.Resampling.BICUBIC
            )

        crop, _ = _enhance_for_ocr(
            crop
        )

        crop_path = os.path.join(
            DEBUG_OUTPUT_DIR,
            "local_ocr",
            (
                f"cell_{attempt_index:02d}_"
                f"r{candidate['period']:02d}_"
                f"d{day_index + 1}.png"
            )
        )

        try:
            crop.save(
                crop_path
            )

            local_items, _ = run_ocr(
                ocr,
                crop_path
            )
        except Exception:
            continue

        mapped_items = []

        for local_item in local_items:

            text = normalize_text(
                local_item.get(
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

            if is_obvious_non_course(
                compact
            ):
                continue

            if compact in {
                "午",
                "休",
                "晨",
                "读"
            }:
                continue

            local_box = local_item[
                "box"
            ]

            box = [
                local_box[0] / scale + crop_left,
                local_box[1] / scale + crop_top,
                local_box[2] / scale + crop_left,
                local_box[3] / scale + crop_top
            ]

            mapped_items.append({
                **local_item,

                "box":
                    box,

                "cx":
                    (
                        box[0]
                        +
                        box[2]
                    ) / 2,

                "cy":
                    (
                        box[1]
                        +
                        box[3]
                    ) / 2,

                "period_number":
                    candidate["period"],

                "period_label":
                    periods[
                        period_index
                    ]["label"],

                "period_y":
                    periods[
                        period_index
                    ]["cy"],

                "period_distance":
                    0.0,

                "day_indices":
                    [day_index],

                "local_ocr":
                    True,

                "local_crop_path":
                    crop_path
            })

        local_text = _cell_raw_text(
            mapped_items
        )

        local_scores = [
            float(
                item.get(
                    "score",
                    0
                )
            )
            for item
            in mapped_items
        ]

        local_score = (
            sum(local_scores)
            /
            len(local_scores)
            if local_scores
            else 0.0
        )

        current_text = candidate[
            "current_text"
        ]

        current_score = candidate[
            "current_score"
        ]

        current_len = len(
            current_text
        )

        local_len = len(
            local_text
        )

        # 严格接受规则：
        # 1. 空格子：必须得到 >=2 字且置信度 >= 0.70；
        # 2. 短文本：新结果必须更长；
        # 3. 一般低置信度：新结果不得明显更短，且置信度不明显下降。
        accept = False

        if not local_text:
            continue

        if not current_text:
            accept = (
                local_len >= 2
                and
                local_score
                >= LOCAL_OCR_EMPTY_MIN_CONFIDENCE
            )

        elif current_len <= LOCAL_OCR_SHORT_TEXT_LENGTH:
            accept = (
                local_len > current_len
                and
                local_score
                >= LOCAL_OCR_REPAIR_MIN_CONFIDENCE
            )

        else:
            current_canonical = (
                canonicalize_course_text(
                    current_text
                )
            )

            local_canonical = (
                canonicalize_course_text(
                    local_text
                )
            )

            # 对普通文本，只有明确变长才替换。
            # 对已知 OCR 别名，只有局部 OCR 结果置信度更高时才替换。
            canonical_improved = (
                local_canonical
                !=
                current_canonical
                and
                local_score
                >=
                current_score + 0.03
            )

            accept = (
                (
                    local_len
                    >
                    current_len
                    and
                    local_score
                    >=
                    current_score - 0.05
                )
                or
                canonical_improved
            )

        if not accept:
            continue

        target_period = candidate[
            "period"
        ]

        target_day = day_index

        new_items = []

        for item in refined_items:

            if (
                item.get(
                    "period_number"
                )
                ==
                target_period
                and
                len(
                    item.get(
                        "day_indices",
                        []
                    )
                )
                == 1
                and
                item["day_indices"][0]
                ==
                target_day
            ):
                continue

            new_items.append(
                item
            )

        refined_items = (
            new_items
            +
            mapped_items
        )

        accepted += 1

    elapsed = (
        time.perf_counter()
        - started
    )

    return (
        refined_items,
        {
            "candidates":
                len(candidates),
            "attempted":
                len(selected),
            "accepted":
                accepted,
            "time_seconds":
                round(
                    elapsed,
                    3
                )
        }
    )


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
    items,
    image_path=None,
    ocr=None
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
    #
    # 现在不再把左侧节次标签作为必要条件。
    # 优先从五个星期列里的课程文字 Y 坐标直接建立课程行。
    #

    periods = detect_periods(
        items,
        day_columns,
        image_path
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



        compact_text = re.sub(
            r"\s+",
            "",
            text
        )

        if compact_text in {
            "午",
            "休",
            "晨",
            "读"
        }:
            ignored.append({
                "text":
                    text,
                "reason":
                    "活动名称残片"
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
    # 可疑单元格局部补识别
    # --------------------------------------------------------
    #
    # 在教师过滤前执行，让局部 OCR 结果也参与角色判定。
    #

    local_ocr_start = (
        time.perf_counter()
    )

    valid_items, local_ocr_stats = (
        refine_suspicious_cells(
            valid_items,
            periods,
            day_columns,
            image_path,
            ocr
        )
    )

    local_ocr_time = (
        time.perf_counter()
        -
        local_ocr_start
    )

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

        "local_ocr":
            {
                **local_ocr_stats,
                "time_seconds":
                    round(
                        local_ocr_time,
                        3
                    )
            },

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

    # 低置信度提醒：与程序已有 LOW_CONFIDENCE_THRESHOLD 保持一致。
    # 黄色仅标记“需要人工核查”的课程格，不修改文字内容。
    low_confidence_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC"
    )

    # 表头与单元格采用适合人工校对的基础格式。
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row_index, schedule_row in enumerate(
        schedule,
        start=2
    ):
        # 第一列节次
        ws.cell(
            row=row_index,
            column=1
        ).alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        for day_index, day in enumerate(
            day_columns,
            start=2
        ):
            excel_cell = ws.cell(
                row=row_index,
                column=day_index
            )

            excel_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            course_cell = schedule_row[
                "cells"
            ][
                day["label"]
            ]

            confidence = course_cell.get(
                "min_confidence"
            )

            if (
                confidence is not None
                and
                confidence < LOW_CONFIDENCE_THRESHOLD
                and
                course_cell.get(
                    "text",
                    ""
                )
            ):
                excel_cell.fill = (
                    low_confidence_fill
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
    # V5.3 图像预处理
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
        print(
            f"黑边检测尺寸（左、上、右、下）："
            f"{preprocess_info['detected_border_runs']}"
        )
    else:
        print(
            "未检测到需要裁剪的黑边"
        )

    if preprocess_info["content_cropped"]:
        print(
            f"检测到白边/空白边并自动裁剪："
            f"{preprocess_info['content_crop_box']}"
        )
    else:
        print(
            "未检测到需要裁剪的白边/空白边"
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

    if preprocess_info["ocr_contrast_enhanced"]:
        print(
            "检测到低对比度，已对亮度通道做轻度增强："
            f"{preprocess_info['ocr_contrast_factor']:.2f}x"
        )
    else:
        print(
            "对比度正常：不做额外 OCR 图像增强"
        )

    print(
        "OCR专用清洗："
        f"{preprocess_info['ocr_cleaning_method']}"
    )

    print(
        "OCR对比清洗图："
        f"{preprocess_info['ocr_input_path']}"
    )

    print(
        "PaddleOCR实际输入："
        f"{preprocess_info['ocr_actual_input_path']}"
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

    # 经过本次 DZ048 实测：
    # 二值化/去线版本会把网格和小字一起“切碎”。
    # 因此 PaddleOCR 恢复使用裁边+放大后的自然图像。
    items, res = run_ocr(
        ocr,
        preprocess_info["ocr_actual_input_path"]
    )

    ocr_time = (
        time.perf_counter()
        - ocr_run_start
    )

    parsed = parse_orientation(
        items,
        preprocessed_path,
        ocr
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

    local_ocr_stats = parsed.get(
        "local_ocr",
        {}
    )

    if local_ocr_stats.get(
        "attempted",
        0
    ):
        print(
            "局部补识别："
            f"候选 {local_ocr_stats.get('candidates', 0)}，"
            f"尝试 {local_ocr_stats.get('attempted', 0)}，"
            f"成功 {local_ocr_stats.get('accepted', 0)}，"
            f"耗时 {local_ocr_stats.get('time_seconds', 0):.2f} 秒"
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
