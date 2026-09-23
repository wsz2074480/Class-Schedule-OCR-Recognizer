import os
import time
from paddleocr import PaddleOCR

IMAGE_PATH = "test.png"
OUTPUT_DIR = "output"

os.makedirs(OUTPUT_DIR, exist_ok=True)

print("=" * 60)
print("正在启动 PaddleOCR...")
print("=" * 60)

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

print(f"OCR 初始化完成，用时：{init_time:.2f} 秒")
print()
print("正在识别图片...")
print()

start_ocr = time.perf_counter()

results = ocr.predict(IMAGE_PATH)

ocr_time = time.perf_counter() - start_ocr

print(f"OCR 识别完成，用时：{ocr_time:.2f} 秒")
print()

for res in results:
    # 保存完整 JSON
    res.save_to_json(OUTPUT_DIR)

    # 保存带识别框的图片
    res.save_to_img(OUTPUT_DIR)

    # 获取 JSON 数据
    data = res.json

    # 兼容 PaddleOCR 当前返回结构
    if isinstance(data, dict) and "res" in data:
        data = data["res"]

    texts = data.get("rec_texts", [])
    scores = data.get("rec_scores", [])
    boxes = data.get("rec_boxes", [])

    print("=" * 60)
    print(f"识别到 {len(texts)} 个文本区域")
    print("=" * 60)

    for i, text in enumerate(texts):
        score = scores[i] if i < len(scores) else 0
        box = boxes[i] if i < len(boxes) else []

        print(f"{i + 1:03d} | {text:<20} | 置信度: {score:.3f} | 坐标: {box}")

print()
print("=" * 60)
print("测试完成")
print(f"识别耗时：{ocr_time:.2f} 秒")
print(f"结果目录：{os.path.abspath(OUTPUT_DIR)}")
print("=" * 60)