import json
import os
import sys
import time

from paddleocr import PaddleOCR


# ============================================================
# 使用方法
#
# python debug_ocr.py test.png
#
# 也可以：
#
# python debug_ocr.py complex1.png
# python debug_ocr.py complex2.png
# ============================================================


if len(sys.argv) < 2:
    print("用法：python debug_ocr.py 图片文件名")
    print("例如：python debug_ocr.py test.png")
    sys.exit(1)


IMAGE_PATH = sys.argv[1]

if not os.path.exists(IMAGE_PATH):
    print(f"找不到图片：{IMAGE_PATH}")
    sys.exit(1)


OUTPUT_DIR = "debug_output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# OCR
# ============================================================

print("=" * 72)
print("启动 PaddleOCR")
print("=" * 72)

start = time.perf_counter()

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
    f"初始化完成："
    f"{time.perf_counter() - start:.2f} 秒"
)

print()
print(f"正在识别：{IMAGE_PATH}")

start = time.perf_counter()

results = ocr.predict(
    IMAGE_PATH
)

ocr_time = time.perf_counter() - start

print(
    f"OCR 完成："
    f"{ocr_time:.2f} 秒"
)

print()


# ============================================================
# 获取结果
# ============================================================

if not results:
    print("OCR 没有返回结果")
    sys.exit(1)


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
# 构造详细结果
# ============================================================

items = []

for i, text in enumerate(texts):

    if hasattr(boxes[i], "tolist"):
        box = boxes[i].tolist()
    else:
        box = boxes[i]

    score = (
        float(scores[i])
        if i < len(scores)
        else None
    )

    x1, y1, x2, y2 = box

    cx = (
        float(x1)
        + float(x2)
    ) / 2

    cy = (
        float(y1)
        + float(y2)
    ) / 2

    items.append({
        "index": i,
        "text": str(text),
        "confidence": score,
        "bbox": [
            float(x1),
            float(y1),
            float(x2),
            float(y2)
        ],
        "center": [
            round(cx, 1),
            round(cy, 1)
        ]
    })


# ============================================================
# 按 Y 再按 X 排序
# ============================================================

items_sorted = sorted(
    items,
    key=lambda x: (
        x["center"][1],
        x["center"][0]
    )
)


# ============================================================
# 控制台输出
# ============================================================

print("=" * 120)
print(
    f"共识别到 {len(items_sorted)} 个文字区域"
)
print("=" * 120)

for item in items_sorted:

    print(
        f"{item['index']:03d} | "
        f"{item['text']:<30} | "
        f"置信度={item['confidence']:.3f} | "
        f"中心=({item['center'][0]:.1f}, "
        f"{item['center'][1]:.1f}) | "
        f"bbox={item['bbox']}"
    )


# ============================================================
# 保存 TXT
# ============================================================

base_name = os.path.splitext(
    os.path.basename(IMAGE_PATH)
)[0]


txt_path = os.path.join(
    OUTPUT_DIR,
    f"{base_name}_raw.txt"
)


with open(
    txt_path,
    "w",
    encoding="utf-8-sig"
) as f:

    f.write(
        f"图片：{IMAGE_PATH}\n"
    )

    f.write(
        f"OCR耗时：{ocr_time:.3f} 秒\n"
    )

    f.write(
        f"文字区域：{len(items_sorted)}\n"
    )

    f.write(
        "=" * 120 + "\n"
    )

    for item in items_sorted:

        f.write(
            f"{item['index']:03d} | "
            f"{item['text']} | "
            f"置信度={item['confidence']:.3f} | "
            f"中心={item['center']} | "
            f"bbox={item['bbox']}\n"
        )


# ============================================================
# 保存 JSON
# ============================================================

json_path = os.path.join(
    OUTPUT_DIR,
    f"{base_name}_raw.json"
)


json_data = {
    "image": IMAGE_PATH,
    "ocr_time_seconds":
        round(ocr_time, 3),
    "count":
        len(items_sorted),
    "items":
        items_sorted
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
# 保存可视化结果
# ============================================================

try:

    res.save_to_img(
        save_path=OUTPUT_DIR
    )

    print()
    print(
        "已保存 OCR 标注图片到："
        f"{OUTPUT_DIR}"
    )

except Exception as e:

    print(
        f"保存标注图片失败：{e}"
    )


# ============================================================
# 完成
# ============================================================

print()
print("=" * 72)
print("诊断完成")
print("=" * 72)

print(
    f"原始 TXT：{txt_path}"
)

print(
    f"原始 JSON：{json_path}"
)

print("=" * 72)