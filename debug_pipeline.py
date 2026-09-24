import json
import os
import sys
import time
from collections import OrderedDict

from PIL import Image

from paddleocr import PaddleOCR

import parse_course_v5 as pipeline


# ============================================================
# CourseOCR V5.3 完整诊断程序
#
# 作用：
# 1. 按生产程序的实际流程执行黑边/白边裁剪、放大、低对比度处理；
# 2. 保存每一个预处理阶段，便于人工查看；
# 3. 分别对“原图”和“最终预处理图”各跑一次 OCR（仅用于诊断）；
# 4. 保存两套原始 OCR 结果、标注图；
# 5. 对最终预处理结果再运行一次课程表 Parser；
# 6. 不修改正式生产程序的行为。
#
# 使用：
#   python debug_pipeline.py "图片完整路径"
#
# 输出：
#   debug_output\pipeline_<图片名>\
# ============================================================


OUTPUT_ROOT = "debug_output"


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )


def safe_float(value):
    try:
        return round(float(value), 4)
    except Exception:
        return None


def collect_ocr_items(res):
    data = res.json

    if (
        isinstance(data, dict)
        and "res" in data
    ):
        data = data["res"]

    texts = data.get("rec_texts", [])
    scores = data.get("rec_scores", [])
    boxes = data.get("rec_boxes", [])

    items = []

    count = min(
        len(texts),
        len(boxes)
    )

    for i in range(count):
        text = str(texts[i]).strip()

        if hasattr(boxes[i], "tolist"):
            box = boxes[i].tolist()
        else:
            box = list(boxes[i])

        x1, y1, x2, y2 = [
            float(value)
            for value in box
        ]

        items.append({
            "index": i,
            "text": text,
            "confidence":
                safe_float(
                    scores[i]
                    if i < len(scores)
                    else None
                ),
            "bbox": [
                round(x1, 1),
                round(y1, 1),
                round(x2, 1),
                round(y2, 1)
            ],
            "center": [
                round((x1 + x2) / 2, 1),
                round((y1 + y2) / 2, 1)
            ]
        })

    items.sort(
        key=lambda item: (
            item["center"][1],
            item["center"][0]
        )
    )

    return data, items


def run_ocr_once(ocr, image_path):
    start = time.perf_counter()

    results = ocr.predict(
        image_path
    )

    elapsed = (
        time.perf_counter()
        - start
    )

    if not results:
        raise RuntimeError(
            f"OCR 没有返回结果：{image_path}"
        )

    data, items = collect_ocr_items(
        results[0]
    )

    return {
        "elapsed": round(
            elapsed,
            3
        ),
        "data": data,
        "items": items,
        "result": results[0]
    }


def save_stage_images(
    original,
    output_dir
):
    """
    与正式 V5.3 完全相同的预处理步骤，但把每一步分别保存下来。
    """

    stage_info = OrderedDict()

    cropped, crop_info = (
        pipeline._find_dark_border_crop(
            original
        )
    )

    stage1 = os.path.join(
        output_dir,
        "01_after_black_border_crop.png"
    )

    cropped.save(stage1)

    stage_info["black_border_crop"] = {
        **crop_info,
        "path": stage1,
        "size": list(cropped.size)
    }

    content_cropped, content_info = (
        pipeline._find_content_crop(
            cropped
        )
    )

    stage2 = os.path.join(
        output_dir,
        "02_after_white_border_crop.png"
    )

    content_cropped.save(stage2)

    stage_info["white_border_crop"] = {
        **content_info,
        "path": stage2,
        "size": list(content_cropped.size)
    }

    resized, resize_info = (
        pipeline._upscale_if_needed(
            content_cropped
        )
    )

    stage3 = os.path.join(
        output_dir,
        "03_after_upscale.png"
    )

    resized.save(stage3)

    stage_info["upscale"] = {
        **resize_info,
        "path": stage3,
        "size": list(resized.size)
    }

    processed, enhance_info = (
        pipeline._enhance_for_ocr(
            resized
        )
    )

    stage4 = os.path.join(
        output_dir,
        "04_final_ocr_input.png"
    )

    processed.save(stage4)

    stage_info["enhancement"] = {
        **enhance_info,
        "path": stage4,
        "size": list(processed.size)
    }

    return (
        processed,
        stage_info
    )


def write_text_report(
    path,
    image_path,
    original_result,
    processed_result,
    parsed,
    stage_info
):
    lines = []

    lines.append("=" * 100)
    lines.append("CourseOCR V5.3 完整诊断")
    lines.append("=" * 100)
    lines.append("")
    lines.append(f"输入图片：{image_path}")
    lines.append(
        f"原图尺寸："
        f"{original_result['image_size'][0]}x"
        f"{original_result['image_size'][1]}"
    )
    lines.append(
        f"预处理图尺寸："
        f"{processed_result['image_size'][0]}x"
        f"{processed_result['image_size'][1]}"
    )
    lines.append("")

    lines.append("-" * 100)
    lines.append("一、预处理阶段")
    lines.append("-" * 100)

    for name, info in stage_info.items():
        lines.append(
            f"{name}: {json.dumps(info, ensure_ascii=False)}"
        )

    lines.append("")
    lines.append("-" * 100)
    lines.append("二、原图 OCR")
    lines.append("-" * 100)
    lines.append(
        f"耗时：{original_result['elapsed']:.3f} 秒"
    )
    lines.append(
        f"文字区域：{len(original_result['items'])}"
    )

    for item in original_result["items"]:
        lines.append(
            f"{item['index']:03d} | "
            f"{item['text']} | "
            f"score={item['confidence']} | "
            f"center={item['center']} | "
            f"bbox={item['bbox']}"
        )

    lines.append("")
    lines.append("-" * 100)
    lines.append("三、最终预处理图 OCR")
    lines.append("-" * 100)
    lines.append(
        f"耗时：{processed_result['elapsed']:.3f} 秒"
    )
    lines.append(
        f"文字区域：{len(processed_result['items'])}"
    )

    for item in processed_result["items"]:
        lines.append(
            f"{item['index']:03d} | "
            f"{item['text']} | "
            f"score={item['confidence']} | "
            f"center={item['center']} | "
            f"bbox={item['bbox']}"
        )

    lines.append("")
    lines.append("-" * 100)
    lines.append("四、课程表 Parser")
    lines.append("-" * 100)

    lines.append(
        f"success={parsed.get('success')}"
    )
    lines.append(
        f"reason={parsed.get('reason')}"
    )
    lines.append(
        f"星期列={len(parsed.get('day_columns', []))}"
    )
    lines.append(
        f"节次={len(parsed.get('periods', []))}"
    )
    lines.append(
        f"有效课程格={parsed.get('non_empty_cells', 0)}"
    )
    lines.append(
        f"结构分数={parsed.get('structure_score', 0):.2f}"
    )
    lines.append(
        f"教师姓名候选={parsed.get('teacher_names', [])}"
    )

    if parsed.get("schedule"):
        lines.append("")
        lines.append("最终课程表：")

        day_columns = parsed.get(
            "day_columns",
            []
        )

        headers = (
            ["节次"]
            + [
                day["label"]
                for day in day_columns
            ]
        )

        lines.append(
            "\t".join(headers)
        )

        for row in parsed["schedule"]:
            values = [
                row["period_label"]
            ]

            for day in day_columns:
                values.append(
                    row["cells"][
                        day["label"]
                    ]["text"]
                )

            lines.append(
                "\t".join(values)
            )

    lines.append("")
    lines.append("=" * 100)

    with open(
        path,
        "w",
        encoding="utf-8-sig"
    ) as f:
        f.write(
            "\n".join(lines)
        )


def save_ocr_result(
    output_dir,
    prefix,
    result
):
    save_json(
        os.path.join(
            output_dir,
            f"{prefix}_raw.json"
        ),
        {
            "elapsed": result["elapsed"],
            "image_size": result["image_size"],
            "count": len(result["items"]),
            "items": result["items"]
        }
    )

    txt_path = os.path.join(
        output_dir,
        f"{prefix}_raw.txt"
    )

    with open(
        txt_path,
        "w",
        encoding="utf-8-sig"
    ) as f:
        f.write(
            f"图片：{result['image_path']}\n"
        )
        f.write(
            f"尺寸：{result['image_size']}\n"
        )
        f.write(
            f"OCR耗时：{result['elapsed']:.3f} 秒\n"
        )
        f.write(
            f"文字区域：{len(result['items'])}\n"
        )
        f.write(
            "=" * 120
            + "\n"
        )

        for item in result["items"]:
            f.write(
                f"{item['index']:03d} | "
                f"{item['text']} | "
                f"置信度={item['confidence']} | "
                f"中心={item['center']} | "
                f"bbox={item['bbox']}\n"
            )


def main():
    if len(sys.argv) < 2:
        print(
            '用法：python debug_pipeline.py "图片完整路径"'
        )
        sys.exit(1)

    image_path = os.path.abspath(
        sys.argv[1]
    )

    if not os.path.exists(image_path):
        print(
            f"找不到图片：{image_path}"
        )
        sys.exit(1)

    base_name = os.path.splitext(
        os.path.basename(image_path)
    )[0]

    output_dir = os.path.join(
        OUTPUT_ROOT,
        f"pipeline_{base_name}"
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    print("=" * 80)
    print("CourseOCR V5.3 完整诊断")
    print("=" * 80)
    print(
        f"输入：{image_path}"
    )

    original = Image.open(
        image_path
    ).convert("RGB")

    print(
        f"原图尺寸：{original.size[0]}x{original.size[1]}"
    )

    # --------------------------------------------------------
    # 保存各个预处理阶段
    # --------------------------------------------------------

    stage_start = time.perf_counter()

    processed, stage_info = (
        save_stage_images(
            original,
            output_dir
        )
    )

    stage_elapsed = (
        time.perf_counter()
        - stage_start
    )

    print(
        f"预处理阶段保存完成："
        f"{stage_elapsed:.3f} 秒"
    )

    # --------------------------------------------------------
    # 加载模型
    # --------------------------------------------------------

    print()
    print("正在加载 PaddleOCR...")

    model_start = time.perf_counter()

    ocr = PaddleOCR(
        text_detection_model_name="PP-OCRv5_mobile_det",
        text_recognition_model_name="PP-OCRv5_mobile_rec",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        device="cpu",
        engine="paddle"
    )

    model_elapsed = (
        time.perf_counter()
        - model_start
    )

    print(
        f"OCR 模型加载完成："
        f"{model_elapsed:.3f} 秒"
    )

    # --------------------------------------------------------
    # 原图 OCR
    # --------------------------------------------------------

    print()
    print("正在执行原图 OCR...")

    original_ocr = run_ocr_once(
        ocr,
        image_path
    )

    original_result = {
        **original_ocr,
        "image_path": image_path,
        "image_size": list(
            original.size
        )
    }

    print(
        f"原图 OCR："
        f"{len(original_result['items'])} 个文字区域，"
        f"{original_result['elapsed']:.3f} 秒"
    )

    try:
        original_ocr["result"].save_to_img(
            save_path=output_dir
        )
    except Exception as exc:
        print(
            f"原图 OCR 标注图保存失败：{exc}"
        )

    save_ocr_result(
        output_dir,
        "original",
        original_result
    )

    # --------------------------------------------------------
    # 最终预处理图 OCR
    # --------------------------------------------------------

    final_path = stage_info[
        "enhancement"
    ]["path"]

    print()
    print("正在执行最终预处理图 OCR...")

    processed_ocr = run_ocr_once(
        ocr,
        final_path
    )

    processed_result = {
        **processed_ocr,
        "image_path": final_path,
        "image_size": list(
            processed.size
        )
    }

    print(
        f"预处理图 OCR："
        f"{len(processed_result['items'])} 个文字区域，"
        f"{processed_result['elapsed']:.3f} 秒"
    )

    try:
        processed_ocr["result"].save_to_img(
            save_path=output_dir
        )
    except Exception as exc:
        print(
            f"预处理图 OCR 标注图保存失败：{exc}"
        )

    save_ocr_result(
        output_dir,
        "processed",
        processed_result
    )

    # --------------------------------------------------------
    # Parser
    # --------------------------------------------------------

    items_for_parser = []

    for item in processed_result["items"]:
        items_for_parser.append({
            "index": item["index"],
            "text": item["text"],
            "score": (
                float(item["confidence"])
                if item["confidence"] is not None
                else 0.0
            ),
            "box": item["bbox"],
            "cx": item["center"][0],
            "cy": item["center"][1]
        })

    parsed = pipeline.parse_orientation(
        items_for_parser
    )

    report_path = os.path.join(
        output_dir,
        "diagnosis.txt"
    )

    write_text_report(
        report_path,
        image_path,
        original_result,
        processed_result,
        parsed,
        stage_info
    )

    save_json(
        os.path.join(
            output_dir,
            "diagnosis.json"
        ),
        {
            "image": image_path,
            "model_load_seconds":
                round(
                    model_elapsed,
                    3
                ),
            "preprocess_seconds":
                round(
                    stage_elapsed,
                    3
                ),
            "original": {
                "image_size":
                    original_result["image_size"],
                "ocr_seconds":
                    original_result["elapsed"],
                "count":
                    len(original_result["items"])
            },
            "processed": {
                "image_size":
                    processed_result["image_size"],
                "ocr_seconds":
                    processed_result["elapsed"],
                "count":
                    len(processed_result["items"])
            },
            "preprocess_stages":
                stage_info,
            "parser": parsed
        }
    )

    print()
    print("=" * 80)
    print("诊断完成")
    print("=" * 80)
    print(
        f"原图 OCR："
        f"{len(original_result['items'])} 个文字区域"
    )
    print(
        f"预处理图 OCR："
        f"{len(processed_result['items'])} 个文字区域"
    )
    print(
        f"Parser："
        f"{parsed.get('success')} / "
        f"{parsed.get('reason')}"
    )
    print()
    print(
        f"诊断目录：{output_dir}"
    )
    print(
        f"详细报告：{report_path}"
    )
    print("=" * 80)


if __name__ == "__main__":
    main()
