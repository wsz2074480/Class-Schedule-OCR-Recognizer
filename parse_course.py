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

# 两个 OCR 文字中心点在这个距离以内，认为属于同一课程行
# 当前这张课程表约 40px，设置 55 比较保险
ROW_CLUSTER_MAX_DISTANCE = 55

# 单元格低置信度阈值
LOW_CONFIDENCE_THRESHOLD = 0.90


# ============================================================
# 工具函数
# ============================================================

def normalize_text(text):
    """清理 OCR 文本"""
    if text is None:
        return ""

    text = unicodedata.normalize("NFKC", str(text))
    text = text.strip()

    # 去掉多余空白，但保留英文单词之间的空格
    text = re.sub(r"[ \t]+", " ", text)

    return text


def box_to_list(box):
    """numpy / list → 普通 list"""
    if hasattr(box, "tolist"):
        box = box.tolist()

    return [float(x) for x in box]


def get_box_center(box):
    x1, y1, x2, y2 = box

    cx = (x1 + x2) / 2
    cy = (y1 + y2) / 2

    return cx, cy


def normalize_weekday(text):
    """
    星期一 / 周一 / 星期1 / 周1
    → 周一
    """

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
        return mapping.get(m.group(1))

    return None


def parse_period_label(text):
    """
    识别：

    1
    2
    第1节
    第2课
    上午1
    下午2
    3-4
    第3-4节

    返回：
        {
            "start": 3,
            "label": "3-4"
        }

    识别不到则返回 None
    """

    text = normalize_text(text)

    # 去除时间段名称
    cleaned = re.sub(
        r"^(上午|下午|早上|晚上|晚自习)\s*",
        "",
        text
    )

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


def join_cell_text(items):
    """
    将同一个单元格中的多个 OCR 区域合并。

    例如：
        地方课程
        （快乐英语）

    → 地方课程（快乐英语）
    """

    if not items:
        return ""

    items = sorted(
        items,
        key=lambda x: (x["cy"], x["cx"])
    )

    parts = []

    for item in items:
        text = item["text"]

        if not text:
            continue

        parts.append(text)

    if not parts:
        return ""

    # 中文场景直接连接
    result = ""

    for part in parts:

        if not result:
            result = part
            continue

        # 如果前后明显是英文，则加一个空格
        if (
            re.search(r"[A-Za-z0-9]$", result)
            and re.match(r"^[A-Za-z0-9]", part)
        ):
            result += " " + part
        else:
            result += part

    return result


def cluster_rows(items):
    """
    根据 OCR 文字中心 Y 坐标自动聚类课程行。
    """

    if not items:
        return []

    items = sorted(
        items,
        key=lambda x: x["cy"]
    )

    # 所有 Y 差值
    gaps = []

    for i in range(1, len(items)):
        gap = items[i]["cy"] - items[i - 1]["cy"]

        if gap > 2:
            gaps.append(gap)

    if gaps:
        # 课程行间距通常远大于同一单元格的多行文字间距
        typical_gap = median(gaps)

        threshold = max(
            ROW_CLUSTER_MAX_DISTANCE,
            min(80, typical_gap * 0.45)
        )
    else:
        threshold = ROW_CLUSTER_MAX_DISTANCE

    groups = []
    current = [items[0]]

    for item in items[1:]:

        current_y = median(
            x["cy"] for x in current
        )

        if abs(item["cy"] - current_y) <= threshold:
            current.append(item)
        else:
            groups.append(current)
            current = [item]

    groups.append(current)

    return groups


def find_nearest_day(cx, day_centers):
    """根据 X 坐标找到最近的星期列"""

    return min(
        day_centers.keys(),
        key=lambda day: abs(cx - day_centers[day])
    )


# ============================================================
# 启动 OCR
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

init_time = time.perf_counter() - start_init

print(f"OCR 初始化完成：{init_time:.2f} 秒")
print()


# ============================================================
# OCR
# ============================================================

print("正在识别图片...")

start_ocr = time.perf_counter()

results = ocr.predict(IMAGE_PATH)

ocr_time = time.perf_counter() - start_ocr

print(f"OCR 识别完成：{ocr_time:.2f} 秒")
print()


if not results:
    raise RuntimeError("OCR 没有返回结果")


res = results[0]

data = res.json

if isinstance(data, dict) and "res" in data:
    data = data["res"]


texts = data.get("rec_texts", [])
scores = data.get("rec_scores", [])
boxes = data.get("rec_boxes", [])


# ============================================================
# 整理 OCR 数据
# ============================================================

items = []

for i, text in enumerate(texts):

    text = normalize_text(text)

    if not text:
        continue

    box = box_to_list(boxes[i])

    cx, cy = get_box_center(box)

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


print(f"OCR 文本区域：{len(items)} 个")
print()


# ============================================================
# 识别星期标题
# ============================================================

week_items = []

for item in items:

    day = normalize_weekday(item["text"])

    if day:

        item["day"] = day

        week_items.append(item)


if not week_items:
    raise RuntimeError(
        "没有识别到星期标题。"
    )


week_items.sort(
    key=lambda x: x["cx"]
)


# 如果 OCR 重复识别同一个星期，只保留一个
day_centers = {}

for item in week_items:

    day = item["day"]

    if day not in day_centers:
        day_centers[day] = item["cx"]


days = [
    item["day"]
    for item in week_items
    if item["day"] in day_centers
]

# 去重
days = list(dict.fromkeys(days))

day_centers = {
    day: day_centers[day]
    for day in days
}


print("=" * 72)
print("识别到的星期")
print("=" * 72)

for day in days:

    print(
        f"{day:<4} "
        f"X={day_centers[day]:.1f}"
    )

print()


# ============================================================
# 计算星期列间距
# ============================================================

day_x_sorted = sorted(
    day_centers.values()
)

if len(day_x_sorted) >= 2:

    day_gaps = [
        day_x_sorted[i] - day_x_sorted[i - 1]
        for i in range(1, len(day_x_sorted))
    ]

    column_width = median(day_gaps)

else:

    column_width = 160


# 左边节次编号区域的右边界
period_left_boundary = (
    day_x_sorted[0]
    - column_width * 0.55
)


# ============================================================
# 找节次编号
# ============================================================

header_y = median(
    item["cy"]
    for item in week_items
)


period_candidates = []

for item in items:

    # 已经确定是星期标题
    if normalize_weekday(item["text"]):
        continue

    # 必须在星期列左侧
    if item["cx"] >= period_left_boundary:
        continue

    # 不处理标题区域
    if item["cy"] <= header_y:
        continue

    parsed = parse_period_label(
        item["text"]
    )

    if parsed:

        period_candidates.append({
            **item,
            **parsed
        })


print("=" * 72)
print("识别到的节次编号")
print("=" * 72)

for item in sorted(
    period_candidates,
    key=lambda x: x["cy"]
):

    print(
        f"文字={item['text']:<8} "
        f"节次={item['label']:<5} "
        f"Y={item['cy']:.1f} "
        f"置信度={item['score']:.3f}"
    )

print()


# ============================================================
# 筛选课程文字
# ============================================================

course_items = []

for item in items:

    text = item["text"]

    # 星期标题
    if normalize_weekday(text):
        continue

    # 节次编号
    if parse_period_label(text):
        continue

    # 标题区域
    if item["cy"] <= header_y + max(
        40,
        column_width * 0.30
    ):
        continue

    # 左侧节次区域
    if item["cx"] < period_left_boundary:
        continue

    course_items.append(item)


if not course_items:
    raise RuntimeError(
        "没有找到课程文字。"
    )


# ============================================================
# 按 Y 聚类课程行
# ============================================================

row_groups = cluster_rows(
    course_items
)

row_groups.sort(
    key=lambda group: median(
        item["cy"]
        for item in group
    )
)


print("=" * 72)
print(f"识别到 {len(row_groups)} 个课程行")
print("=" * 72)
print()


# ============================================================
# 建立课程行
# ============================================================

schedule_rows = []


for row_index, group in enumerate(row_groups):

    row_y = median(
        item["cy"]
        for item in group
    )

    # --------------------------------------------------------
    # 找这一行附近的节次编号
    # --------------------------------------------------------

    nearest_period = None

    if period_candidates:

        nearest_period = min(
            period_candidates,
            key=lambda x: abs(
                x["cy"] - row_y
            )
        )

        # 如果距离太远，不认为它属于这一行
        if abs(
            nearest_period["cy"] - row_y
        ) > ROW_CLUSTER_MAX_DISTANCE * 1.5:

            nearest_period = None


    row = {
        "row_index": row_index,
        "y": row_y,
        "period": None,
        "cells": {}
    }


    # 如果找到了真实节次编号
    if nearest_period:
        row["period"] = nearest_period["label"]


    # --------------------------------------------------------
    # 将课程放入最近的星期列
    # --------------------------------------------------------

    for item in group:

        day = find_nearest_day(
            item["cx"],
            day_centers
        )

        if day not in row["cells"]:
            row["cells"][day] = []

        row["cells"][day].append(item)


    schedule_rows.append(row)


# ============================================================
# 补全没有 OCR 到的节次编号
# ============================================================

# 先记录已经知道的数字节次

known_periods = {}

for i, row in enumerate(schedule_rows):

    if row["period"]:

        parsed = parse_period_label(
            row["period"]
        )

        if parsed:

            known_periods[i] = parsed["start"]


# 对没有识别到节次的行进行推断
for i, row in enumerate(schedule_rows):

    if row["period"]:
        continue

    inferred = None

    # 向前找最近一个已知节次
    previous = None

    for j in range(i - 1, -1, -1):

        if j in known_periods:

            previous = (
                j,
                known_periods[j]
            )

            break

    # 向后找最近一个已知节次
    following = None

    for j in range(i + 1, len(schedule_rows)):

        if j in known_periods:

            following = (
                j,
                known_periods[j]
            )

            break


    # 例如：
    # 第1行没有数字
    # 第2行 OCR 到 2
    # → 第1行推断为 1
    if previous:

        previous_index, previous_number = previous

        inferred = previous_number + (
            i - previous_index
        )

    elif following:

        following_index, following_number = following

        inferred = following_number - (
            following_index - i
        )

    else:

        inferred = i + 1


    row["period"] = str(inferred)


# ============================================================
# 合并单元格 OCR 文字
# ============================================================

for row in schedule_rows:

    for day in days:

        cell_items = row["cells"].get(
            day,
            []
        )

        if not cell_items:
            continue

        text = join_cell_text(
            cell_items
        )

        scores = [
            float(x["score"])
            for x in cell_items
        ]

        # 最低置信度
        min_score = min(scores)

        # 平均置信度
        avg_score = sum(scores) / len(scores)

        # 计算整体 bbox
        x1 = min(
            x["box"][0]
            for x in cell_items
        )

        y1 = min(
            x["box"][1]
            for x in cell_items
        )

        x2 = max(
            x["box"][2]
            for x in cell_items
        )

        y2 = max(
            x["box"][3]
            for x in cell_items
        )

        row["cells"][day] = {
            "text": text,
            "min_confidence": round(
                min_score,
                4
            ),
            "avg_confidence": round(
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
                    "text": x["text"],
                    "confidence": round(
                        float(x["score"]),
                        4
                    ),
                    "bbox": [
                        round(v, 1)
                        for v in x["box"]
                    ]
                }
                for x in cell_items
            ]
        }


# ============================================================
# 输出：打印标准课程表
# ============================================================

print()
print("=" * 100)
print("第二版标准课程表")
print("=" * 100)

print(
    "\t".join(
        ["节次"] + days
    )
)


for row in schedule_rows:

    values = [
        f"第{row['period']}节"
    ]

    for day in days:

        cell = row["cells"].get(
            day
        )

        if isinstance(cell, dict):
            values.append(
                cell["text"]
            )
        else:
            values.append("")

    print(
        "\t".join(values)
    )


# ============================================================
# 输出 CSV
# ============================================================

os.makedirs(
    OUTPUT_DIR,
    exist_ok=True
)

csv_path = os.path.join(
    OUTPUT_DIR,
    "course_schedule_v2.csv"
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
                f"第{row['period']}节"
            ]
            + [
                (
                    row["cells"][day]["text"]
                    if isinstance(
                        row["cells"].get(day),
                        dict
                    )
                    else ""
                )
                for day in days
            ]
        )


# ============================================================
# 输出 JSON
# ============================================================

json_path = os.path.join(
    OUTPUT_DIR,
    "course_schedule_v2.json"
)


json_data = {
    "source_image": IMAGE_PATH,
    "ocr_time_seconds": round(
        ocr_time,
        3
    ),
    "days": days,
    "schedule": []
}


for row in schedule_rows:

    cells = {}

    for day in days:

        cell = row["cells"].get(day)

        if isinstance(cell, dict):
            cells[day] = cell
        else:
            cells[day] = {
                "text": "",
                "min_confidence": None,
                "avg_confidence": None,
                "bbox": None,
                "ocr_items": []
            }

    json_data["schedule"].append({
        "period": row["period"],
        "y": round(
            row["y"],
            1
        ),
        "cells": cells
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
# 输出低置信度单元格
# ============================================================

low_confidence = []

for row in schedule_rows:

    for day in days:

        cell = row["cells"].get(day)

        if not isinstance(cell, dict):
            continue

        if (
            cell["min_confidence"]
            < LOW_CONFIDENCE_THRESHOLD
        ):

            low_confidence.append({
                "period": row["period"],
                "day": day,
                "text": cell["text"],
                "confidence": cell[
                    "min_confidence"
                ]
            })


low_path = os.path.join(
    OUTPUT_DIR,
    "low_confidence_cells.csv"
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

print(f"OCR 耗时：{ocr_time:.2f} 秒")
print(f"课程行数：{len(schedule_rows)}")
print(f"星期数量：{len(days)}")
print()
print(f"CSV：{csv_path}")
print(f"JSON：{json_path}")
print(
    f"低置信度单元格：{low_path}"
)
print(
    f"低置信度数量：{len(low_confidence)}"
)
print("=" * 72)