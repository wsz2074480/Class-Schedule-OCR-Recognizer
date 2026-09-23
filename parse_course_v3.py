import csv
import json
import os
import re
import time
import unicodedata
from statistics import median

from paddleocr import PaddleOCR


# ============================================================
# 配置
# ============================================================

IMAGE_PATH = "test.png"
OUTPUT_DIR = "output"

# 同一课程行内部，多个 OCR 文字的 Y 距离
ROW_CLUSTER_MAX_DISTANCE = 45

# 节次编号与课程行之间最大允许距离
MAX_PERIOD_ROW_DISTANCE = 75

# 低置信度阈值
LOW_CONFIDENCE_THRESHOLD = 0.90


# ============================================================
# 非课程信息
# ============================================================

# 完全匹配时直接排除
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
}


# 出现这些关键词，通常属于备注/说明
NON_COURSE_KEYWORDS = [
    "备注",
    "注：",
    "注:",
    "说明",
    "每学期",
    "每周安排",
    "课程安排说明",
]


# ============================================================
# 基础函数
# ============================================================

def normalize_text(text):
    """统一 OCR 文本格式"""

    if text is None:
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(text)
    )

    text = text.strip()

    # 多个空格合并
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


def get_box_center(box):

    x1, y1, x2, y2 = box

    return (
        (x1 + x2) / 2,
        (y1 + y2) / 2
    )


# ============================================================
# 星期识别
# ============================================================

def normalize_weekday(text):

    text = normalize_text(text)

    mapping = {
        "一": "周一",
        "二": "周二",
        "三": "周三",
        "四": "周四",
        "五": "周五",
        "六": "周六",
        "日": "周日",
        "天": "周日",

        "1": "周一",
        "2": "周二",
        "3": "周三",
        "4": "周四",
        "5": "周五",
        "6": "周六",
        "7": "周日",
    }

    m = re.search(
        r"(?:星期|周)\s*([一二三四五六日天1-7])",
        text
    )

    if m:
        return mapping.get(
            m.group(1)
        )

    return None


# ============================================================
# 节次识别
# ============================================================

def parse_period_label(text):

    text = normalize_text(text)

    cleaned = re.sub(
        r"^(上午|下午|早上|晚上|晚自习)\s*",
        "",
        text
    )

    # 支持：
    # 1
    # 2
    # 第1节
    # 第2课
    # 3-4
    # 第3-4节
    pattern = re.compile(
        r"^(?:第\s*)?"
        r"([0-9]{1,2})"
        r"(?:\s*[-~～—]\s*([0-9]{1,2}))?"
        r"\s*(?:节|课)?$"
    )

    m = pattern.match(cleaned)

    if not m:
        return None

    start = int(m.group(1))

    end = m.group(2)

    if end:
        label = f"{start}-{int(end)}"
    else:
        label = str(start)

    return {
        "start": start,
        "label": label
    }


# ============================================================
# 判断是否为非课程信息
# ============================================================

def is_non_course_text(text):

    text = normalize_text(text)

    if not text:
        return True

    # 完全匹配
    if text in NON_COURSE_EXACT:
        return True

    # 关键词匹配
    for keyword in NON_COURSE_KEYWORDS:
        if keyword in text:
            return True

    return False


def row_is_non_course(cell_texts):

    """
    对整行进行判断。

    例如：
        周一：课间
        周二：操

    合并后：

        课间操

    也可以过滤。
    """

    useful_texts = [
        normalize_text(x)
        for x in cell_texts
        if normalize_text(x)
    ]

    if not useful_texts:
        return True

    whole_text = "".join(
        useful_texts
    )

    # 例如 OCR 被拆成：
    # 课间 + 操
    if "课间操" in whole_text:
        return True

    if "大课间" in whole_text:
        return True

    if "午休" in whole_text:
        return True

    if "午睡" in whole_text:
        return True

    if "眼保健操" in whole_text:
        return True

    if "广播操" in whole_text:
        return True

    if "备注" in whole_text:
        return True

    if "说明" in whole_text:
        return True

    if "每学期" in whole_text:
        return True

    if "每周安排" in whole_text:
        return True

    # 单独文本完全匹配
    if all(
        is_non_course_text(x)
        for x in useful_texts
    ):
        return True

    return False


# ============================================================
# 同一单元格多段 OCR 合并
# ============================================================

def join_cell_text(items):

    if not items:
        return ""

    items = sorted(
        items,
        key=lambda x: (
            x["cy"],
            x["cx"]
        )
    )

    result = ""

    for item in items:

        text = item["text"]

        if not text:
            continue

        if not result:
            result = text
            continue

        # 英文之间增加空格
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
# Y 坐标聚类
# ============================================================

def cluster_rows(items):

    if not items:
        return []

    items = sorted(
        items,
        key=lambda x: x["cy"]
    )

    groups = []

    current = [items[0]]

    for item in items[1:]:

        current_y = median(
            x["cy"]
            for x in current
        )

        if (
            abs(
                item["cy"] - current_y
            )
            <= ROW_CLUSTER_MAX_DISTANCE
        ):
            current.append(item)

        else:
            groups.append(current)
            current = [item]

    groups.append(current)

    return groups


# ============================================================
# 找最近的星期
# ============================================================

def find_nearest_day(
    cx,
    day_centers
):

    return min(
        day_centers.keys(),
        key=lambda day:
        abs(
            cx - day_centers[day]
        )
    )


# ============================================================
# 启动 PaddleOCR
# ============================================================

print("=" * 72)
print("启动 PaddleOCR")
print("=" * 72)

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

init_time = (
    time.perf_counter()
    - start_init
)

print(
    f"OCR 初始化完成："
    f"{init_time:.2f} 秒"
)

print()


# ============================================================
# OCR
# ============================================================

print("正在识别图片...")

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
# 整理 OCR
# ============================================================

items = []

for i, text in enumerate(texts):

    text = normalize_text(text)

    if not text:
        continue

    box = box_to_list(
        boxes[i]
    )

    cx, cy = get_box_center(
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
        "cy": cy,
    })


print(
    f"OCR 文本区域："
    f"{len(items)} 个"
)

print()


# ============================================================
# 识别星期
# ============================================================

week_items = []

for item in items:

    day = normalize_weekday(
        item["text"]
    )

    if day:

        week_items.append({
            **item,
            "day": day
        })


if not week_items:

    raise RuntimeError(
        "没有识别到星期标题"
    )


week_items.sort(
    key=lambda x: x["cx"]
)


day_centers = {}

for item in week_items:

    day = item["day"]

    if day not in day_centers:

        day_centers[day] = (
            item["cx"]
        )


days = list(
    day_centers.keys()
)


print("=" * 72)
print("识别到的星期")
print("=" * 72)

for day in days:

    print(
        f"{day:<4}"
        f" X={day_centers[day]:.1f}"
    )

print()


# ============================================================
# 计算第一星期列的位置
# ============================================================

first_day_x = min(
    day_centers.values()
)

sorted_day_x = sorted(
    day_centers.values()
)


if len(sorted_day_x) >= 2:

    gaps = [
        sorted_day_x[i]
        - sorted_day_x[i - 1]
        for i in range(
            1,
            len(sorted_day_x)
        )
    ]

    column_width = median(
        gaps
    )

else:

    column_width = 160


# ------------------------------------------------------------
# 关键修正：
#
# 节次编号一般在“周一”左边一点。
#
# 不再使用：
#     first_day_x - 0.55 * column_width
#
# 而是使用相对保守的边界：
#     周一左侧约 1/4 列宽
#
# 你的图片：
#     周一 ≈ 245
#     第1节 ≈ 174
#     边界 ≈ 198
#
# 可以正常识别。
# ------------------------------------------------------------

period_left_boundary = (
    first_day_x
    - max(
        40,
        column_width * 0.25
    )
)


print(
    f"节次区域右边界："
    f"X={period_left_boundary:.1f}"
)

print()


# ============================================================
# 识别真实节次编号
# ============================================================

header_y = median(
    item["cy"]
    for item in week_items
)


period_candidates = []

for item in items:

    text = item["text"]

    # 星期标题不能算节次
    if normalize_weekday(text):
        continue

    # 必须位于第一星期列左侧
    if (
        item["cx"]
        >= period_left_boundary
    ):
        continue

    # 必须在表头下面
    if (
        item["cy"]
        <= header_y
    ):
        continue

    parsed = parse_period_label(
        text
    )

    if parsed:

        period_candidates.append({
            **item,
            **parsed
        })


# 同一个节次如果识别到多个，只保留最可信的一个
period_best = {}

for item in period_candidates:

    number = item["start"]

    if (
        number not in period_best
        or item["score"]
        > period_best[number]["score"]
    ):

        period_best[number] = item


period_candidates = sorted(
    period_best.values(),
    key=lambda x: (
        x["start"],
        x["cy"]
    )
)


print("=" * 72)
print("识别到的真实节次")
print("=" * 72)

for item in period_candidates:

    print(
        f"第{item['label']}节 "
        f"Y={item['cy']:.1f} "
        f"X={item['cx']:.1f} "
        f"置信度={item['score']:.3f}"
    )

print()


# ============================================================
# 提取候选课程文字
# ============================================================

course_items = []

for item in items:

    text = item["text"]

    # 星期标题
    if normalize_weekday(text):
        continue

    # 只把左边区域的数字当节次
    if (
        item["cx"]
        < period_left_boundary
        and parse_period_label(text)
    ):
        continue

    # 标题上方
    if (
        item["cy"]
        <= header_y
    ):
        continue

    # 左侧节次区域
    if (
        item["cx"]
        < period_left_boundary
    ):
        continue

    course_items.append(
        item
    )


# ============================================================
# 先按照 Y 聚类
# ============================================================

row_groups = cluster_rows(
    course_items
)

row_groups.sort(
    key=lambda group:
    median(
        item["cy"]
        for item in group
    )
)


# ============================================================
# 将 OCR 行转换成“星期 → 文本”
# 同时过滤非课程行
# ============================================================

candidate_rows = []

for row_index, group in enumerate(
    row_groups
):

    row_y = median(
        item["cy"]
        for item in group
    )

    grouped_cells = {}

    for item in group:

        day = find_nearest_day(
            item["cx"],
            day_centers
        )

        if day not in grouped_cells:

            grouped_cells[day] = []

        grouped_cells[day].append(
            item
        )


    # --------------------------------------------
    # 先把整行合并成文字
    # --------------------------------------------

    cell_texts = []

    for day in days:

        cell_items = grouped_cells.get(
            day,
            []
        )

        text = join_cell_text(
            cell_items
        )

        if text:
            cell_texts.append(
                text
            )


    # --------------------------------------------
    # 判断是不是：
    # 课间操
    # 午休
    # 备注
    # 等非课程信息
    # --------------------------------------------

    if row_is_non_course(
        cell_texts
    ):

        print(
            f"过滤非课程行："
            f"{' / '.join(cell_texts)}"
        )

        continue


    candidate_rows.append({
        "index": row_index,
        "y": row_y,
        "group": group,
        "grouped_cells": grouped_cells
    })


print()

print("=" * 72)
print(
    f"过滤后候选课程行："
    f"{len(candidate_rows)} 个"
)
print("=" * 72)

print()


# ============================================================
# 根据真实“第X节”寻找对应课程行
# ============================================================

schedule_rows = []

used_candidate_rows = set()


for period in period_candidates:

    best_row = None
    best_distance = None

    for row in candidate_rows:

        if row["index"] in used_candidate_rows:
            continue

        distance = abs(
            row["y"]
            - period["cy"]
        )

        if (
            distance
            > MAX_PERIOD_ROW_DISTANCE
        ):
            continue

        if (
            best_distance is None
            or distance < best_distance
        ):

            best_distance = distance
            best_row = row


    # 没找到课程行
    # 不人为生成空节次
    if best_row is None:

        print(
            f"警告：第"
            f"{period['start']}节"
            f"没有找到对应课程行"
        )

        schedule_rows.append({
            "period_number":
                period["start"],
            "period_label":
                period["label"],
            "period_y":
                period["cy"],
            "row":
                None
        })

        continue


    used_candidate_rows.add(
        best_row["index"]
    )

    schedule_rows.append({
        "period_number":
            period["start"],
        "period_label":
            period["label"],
        "period_y":
            period["cy"],
        "row":
            best_row
    })


# ============================================================
# 构造最终单元格
# ============================================================

for schedule_row in schedule_rows:

    row = schedule_row["row"]

    cells = {}

    for day in days:

        cells[day] = {
            "text": "",
            "min_confidence": None,
            "avg_confidence": None,
            "bbox": None,
            "ocr_items": []
        }


    if row is not None:

        grouped_cells = (
            row["grouped_cells"]
        )

        for day in days:

            cell_items = grouped_cells.get(
                day,
                []
            )

            if not cell_items:
                continue


            text = join_cell_text(
                cell_items
            )


            scores = [
                float(
                    item["score"]
                )
                for item in cell_items
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
                for item in cell_items
            )

            y1 = min(
                item["box"][1]
                for item in cell_items
            )

            x2 = max(
                item["box"][2]
                for item in cell_items
            )

            y2 = max(
                item["box"][3]
                for item in cell_items
            )


            cells[day] = {
                "text": text,

                "min_confidence":
                    round(
                        min_score,
                        4
                    ),

                "avg_confidence":
                    round(
                        avg_score,
                        4
                    ),

                "bbox": [
                    round(x1, 1),
                    round(y1, 1),
                    round(x2, 1),
                    round(y2, 1)
                ],

                "ocr_items": [
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
                            round(value, 1)
                            for value
                            in item["box"]
                        ]
                    }

                    for item
                    in cell_items
                ]
            }


    schedule_row["cells"] = cells


# ============================================================
# 按节次排序
# ============================================================

schedule_rows.sort(
    key=lambda row:
    row["period_number"]
)


# ============================================================
# 输出标准课程表
# ============================================================

print()

print("=" * 110)
print("V3.1 标准课程表")
print("=" * 110)

print(
    "\t".join(
        ["节次"] + days
    )
)


for row in schedule_rows:

    values = [
        f"第{row['period_number']}节"
    ]

    for day in days:

        values.append(
            row["cells"][day]["text"]
        )

    print(
        "\t".join(values)
    )


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
    "course_schedule_v3_1.csv"
)


with open(
    csv_path,
    "w",
    newline="",
    encoding="utf-8-sig"
) as f:

    writer = csv.writer(f)

    writer.writerow(
        ["节次"] + days
    )

    for row in schedule_rows:

        writer.writerow(
            [
                f"第{row['period_number']}节"
            ]
            + [
                row["cells"][day]["text"]
                for day in days
            ]
        )


# ============================================================
# JSON
# ============================================================

json_path = os.path.join(
    OUTPUT_DIR,
    "course_schedule_v3_1.json"
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
        days,

    "period_count":
        len(schedule_rows),

    "schedule": []
}


for row in schedule_rows:

    json_data[
        "schedule"
    ].append({

        "period":
            row["period_label"],

        "period_number":
            row["period_number"],

        "cells":
            row["cells"]
    })


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
# 低置信度
# ============================================================

low_confidence = []


for row in schedule_rows:

    for day in days:

        cell = row[
            "cells"
        ][day]

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
                    day,

                "text":
                    cell["text"],

                "confidence":
                    cell[
                        "min_confidence"
                    ]
            })


low_path = os.path.join(
    OUTPUT_DIR,
    "low_confidence_cells_v3_1.csv"
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

print("=" * 72)
print("处理完成")
print("=" * 72)

print(
    f"OCR 耗时："
    f"{ocr_time:.2f} 秒"
)

print(
    f"最终有效节次："
    f"{len(schedule_rows)}"
)

print(
    f"星期数量："
    f"{len(days)}"
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

print(
    f"低置信度单元格："
    f"{len(low_confidence)}"
)

print("=" * 72)