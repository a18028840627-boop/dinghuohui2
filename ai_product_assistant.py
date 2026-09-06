from __future__ import annotations

import copy
import hashlib
import json
import multiprocessing
import os
import queue as queue_module
import re
import subprocess
import shutil
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from typing import Any
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    # The bundled macOS Tcl/Tk is not compatible with tkinterdnd2's tkdnd
    # binary. Keep drag-and-drop for Windows builds and use normal Tk on Mac.
    if sys.platform.startswith("win"):
        from tkinterdnd2 import DND_FILES, TkinterDnD
        AppBase = TkinterDnD.Tk
    else:
        raise ImportError
except Exception:
    DND_FILES = None
    AppBase = tk.Tk

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell


IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
SAMPLE_LIBRARY_DIR = Path(r"D:\工具\AI商品资料整理助手_手写标注样本库")
SAMPLE_LIBRARY_FILE = SAMPLE_LIBRARY_DIR / "handwriting_samples.jsonl"
SAMPLE_LIBRARY_IMAGES = SAMPLE_LIBRARY_DIR / "images"


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("：", ":")


def extract_after_label(text: str, labels: list[str]) -> str:
    # OCR may return one field per line or concatenate several handwritten
    # labels into one line. Keep line boundaries first, then stop at the next
    # known label when a line contains multiple fields.
    known_labels = ["货号", "款号", "编号", "型号", "颜色", "色号", "配色", "码数", "鞋码", "尺码", "号码", "码", "鞋底", "材质"]
    for line in str(text or "").splitlines():
        line = re.sub(r"\s+", "", line)
        for label in labels:
            match = re.search(re.escape(label) + r"[:：]?(.+?)(?=" + "|".join(map(re.escape, known_labels)) + r"[:：]|$)", line, re.I)
            if match:
                value = match.group(1).strip(" _-—:：")
                if value:
                    return value
    return ""


def clean_sku(value: str) -> str:
    value = value.strip(" _-—:：")
    value = re.sub(r"^(货号|款号|编号|型号)\s*[:：]?", "", value, flags=re.I)
    return value.strip()


def clean_sizes(value: str) -> str:
    value = value.replace("，", ",").replace("、", ",").replace("至", "-").replace("~", "-")
    nums = re.findall(r"\d+(?:\.\d+)?", value)
    if len(nums) == 1 and nums[0].isdigit() and len(nums[0]) == 4:
        nums = [nums[0][:2], nums[0][2:]]
    if len(nums) == 2 and all(n.isdigit() for n in nums):
        start, end = int(nums[0]), int(nums[1])
        if 20 <= start <= 60 and 20 <= end <= 60 and start <= end <= start + 20:
            return "，".join(str(n) for n in range(start, end + 1))
    if len(nums) >= 2:
        return "，".join(nums)
    return value.strip(" ，,")


def clean_color(value: str) -> str:
    return value.replace("，", "，").strip(" ，,")


def parse_image_filename(path: str | Path) -> dict[str, str]:
    """Read SKU, color, category and style note from image filenames.

    The category marker is 女, 男, 男女 or 童. Text before it is the color,
    while text after it is a style note such as 革 or 网. A separator between
    the SKU and color is optional.
    """
    stem = Path(path).stem.strip()
    match = re.match(
        r"^(?P<sku>[A-Za-z0-9][A-Za-z0-9_-]*?)[\s_-]*(?P<color>[\u4e00-\u9fff].*?)(?P<gender>男女|女|男|童)(?P<style_note>.*)$",
        stem,
    )
    if not match:
        return {"sku": "", "color": "", "gender": "", "style_note": "", "error": "文件名格式应为：款号+颜色+女、男、男女或童（后可带款式说明）"}
    sku = clean_sku(match.group("sku"))
    color = clean_color(match.group("color"))
    style_note = match.group("style_note").strip(" _-—:：")
    if not sku or not color:
        return {"sku": sku, "color": color, "gender": match.group("gender"), "style_note": style_note, "error": "文件名中的款号或颜色为空"}
    return {"sku": sku, "color": color, "gender": match.group("gender"), "style_note": style_note, "error": ""}


def gender_size_value(gender: str) -> str:
    """Return the standard size run for a filename gender marker."""
    if gender == "女":
        return "，".join(str(size) for size in range(35, 41))
    if gender == "男":
        return "，".join(str(size) for size in range(39, 45))
    if gender == "男女":
        return "，".join(str(size) for size in range(35, 45))
    return ""


def parse_tag_text(text: str) -> dict[str, str]:
    sku = extract_after_label(text, ["货号", "款号", "编号", "型号"])
    color = extract_after_label(text, ["颜色", "色号", "配色"])
    size = extract_after_label(text, ["码数", "鞋码", "尺码", "尺码范围", "号码", "码"])

    # Handwritten tags are sometimes read as separate OCR lines without the label.
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    if not sku:
        for line in lines:
            candidate = re.sub(r"[^A-Za-z0-9-]", "", line)
            if re.search(r"[A-Za-z]", candidate) and re.search(r"\d", candidate) and 3 <= len(candidate) <= 20:
                sku = candidate
                break
    if not color:
        for line in lines:
            if any(k in line for k in ["黑", "白", "红", "粉", "蓝", "绿", "灰", "米", "杏", "卡其", "紫", "黄"]):
                color = re.sub(r"^(颜色|色号)\s*[:：]?", "", line).strip()
                if color:
                    break
    if not size:
        for line in lines:
            if len(re.findall(r"\d+", line)) >= 2:
                size = line
                break

    return {
        "sku": clean_sku(sku),
        "color": clean_color(color),
        "size": clean_sizes(size),
        "raw": text.strip(),
    }


def validate_row(row: dict[str, Any]) -> list[str]:
    """Return user-facing validation errors for a recognized row."""
    errors = []
    if not str(row.get("sku", "")).strip():
        errors.append("货号为空")
    if not str(row.get("color", "")).strip():
        errors.append("颜色为空")
    if not str(row.get("size", "")).strip():
        errors.append("码数为空")
    size_text = str(row.get("size", ""))
    if size_text:
        nums = re.findall(r"\d+(?:\.\d+)?", size_text)
        if len(nums) != len(set(nums)):
            errors.append("码数有重复")
    return errors


def analyze_image_quality(image_path: str) -> list[str]:
    warnings = []
    try:
        from PIL import Image, ImageStat, ImageFilter
        with Image.open(image_path) as image:
            width, height = image.size
            if min(width, height) < 500:
                warnings.append("图片尺寸过小")
            gray = image.convert("L")
            mean = ImageStat.Stat(gray).mean[0]
            if mean < 35:
                warnings.append("图片过暗")
            elif mean > 245:
                warnings.append("图片过亮")
            edges = gray.filter(ImageFilter.FIND_EDGES)
            if ImageStat.Stat(edges).mean[0] < 3.5:
                warnings.append("图片可能模糊")
    except Exception:
        warnings.append("图片无法读取")
    return warnings


def classify_failure(errors: list[str], exception_text: str = "") -> str:
    if exception_text:
        text = exception_text.lower()
        if "model" in text or "paddle" in text:
            return "OCR模型或运行环境异常"
        if "timeout" in text:
            return "识别超时"
        return "图片处理异常"
    if "图片无法读取" in errors:
        return "图片无法读取"
    if any("图片" in error for error in errors):
        return "图片质量不佳"
    if "货号为空" in errors and "颜色为空" in errors and "码数为空" in errors:
        return "未找到吊牌或未识别到文字"
    if "货号重复" in errors:
        return "货号重复"
    return "字段缺失或格式异常"


def save_sample_record(source_path: str, sku: str, color: str, size: str) -> None:
    """Save a corrected row and a stable image copy into the handwriting library."""
    if not source_path:
        raise ValueError("当前结果没有关联原图")
    source = Path(source_path)
    SAMPLE_LIBRARY_DIR.mkdir(parents=True, exist_ok=True)
    SAMPLE_LIBRARY_IMAGES.mkdir(parents=True, exist_ok=True)
    library_name = hashlib.sha1(str(source.resolve()).encode("utf-8")).hexdigest()[:12] + "_" + source.name
    library_image = SAMPLE_LIBRARY_IMAGES / library_name
    if source.exists() and not library_image.exists():
        shutil.copy2(source, library_image)
    records = {}
    if SAMPLE_LIBRARY_FILE.exists():
        for line in SAMPLE_LIBRARY_FILE.read_text(encoding="utf-8").splitlines():
            try:
                item = json.loads(line)
                key = item.get("source_image", item.get("image", ""))
                if key:
                    records[key] = item
            except Exception:
                continue
    records[str(source)] = {"image": str(library_image), "source_image": str(source), "sku": sku, "color": color, "size": clean_sizes(size)}
    SAMPLE_LIBRARY_FILE.write_text("\n".join(json.dumps(records[key], ensure_ascii=False) for key in sorted(records)) + "\n", encoding="utf-8")


def detect_tag_candidate(image_path: str) -> str | None:
    """Find a likely light rectangular hangtag for OCR fallback."""
    try:
        import cv2
        from PIL import Image
        image = cv2.imread(image_path)
        if image is None:
            return None
        h, w = image.shape[:2]
        scale = min(1.0, 1200 / max(h, w))
        small = cv2.resize(image, None, fx=scale, fy=scale)
        gray = cv2.cvtColor(small, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (5, 5), 0)
        binary = cv2.threshold(gray, 175, 255, cv2.THRESH_BINARY)[1]
        edges = cv2.Canny(gray, 60, 160)
        mask = cv2.bitwise_or(binary, edges)
        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        candidates = []
        frame_area = small.shape[0] * small.shape[1]
        for contour in contours:
            x, y, cw, ch = cv2.boundingRect(contour)
            rect_area = cw * ch
            if rect_area < frame_area * 0.006 or rect_area > frame_area * 0.55:
                continue
            ratio = max(cw, ch) / max(1, min(cw, ch))
            if ratio > 5.5:
                continue
            fill = cv2.contourArea(contour) / max(1, rect_area)
            if fill < 0.28:
                continue
            # Prefer tag-sized quadrilaterals with a bright interior.
            brightness = float(gray[y:y + ch, x:x + cw].mean()) / 255
            score = (1 - min(1, abs(ratio - 1.8) / 3)) + fill + brightness
            candidates.append((score, x, y, cw, ch))
        if not candidates:
            return None
        _, x, y, cw, ch = max(candidates)
        left = max(0, int((x / scale) - max(100, ch / scale * 0.8)))
        top = max(0, int((y / scale) - max(100, ch / scale * 0.5)))
        right = min(w, int((x + cw) / scale + max(100, ch / scale * 0.8)))
        bottom = min(h, int((y + ch) / scale + max(120, ch / scale * 0.8)))
        with Image.open(image_path) as source:
            crop = source.crop((left, top, right, bottom)).rotate(90, expand=True)
            tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
            tmp.close()
            crop.save(tmp.name, quality=95)
            return tmp.name
    except Exception:
        return None


def run_ocr(image_path: str, preview_dir: str | None = None) -> tuple[str, str | None]:
    os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
    # PaddleX defaults to ~/.paddlex.  That directory can be read-only on
    # managed Windows computers, which prevents even the model lock file from
    # being created.  Paddle Inference on Windows also cannot reliably open
    # model files from a path containing Chinese characters.  Use a short,
    # ASCII-only temp cache, and seed it from models bundled beside the app.
    app_root = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    cache_root = Path(tempfile.gettempdir()) / "AI_Product_Assistant_paddlex"
    cache_root.mkdir(parents=True, exist_ok=True)
    bundled_models = app_root / "_paddlex_cache" / "official_models"
    target_models = cache_root / "official_models"
    if bundled_models.exists():
        shutil.copytree(bundled_models, target_models, dirs_exist_ok=True)
    os.environ.setdefault("PADDLE_PDX_CACHE_HOME", str(cache_root))
    try:
        from paddleocr import PaddleOCR
    except Exception as exc:
        raise RuntimeError("未安装 PaddleOCR。请先安装 requirements.txt 中的依赖。") from exc

    # Lazily created per process/thread to keep startup fast.
    if not hasattr(run_ocr, "engine"):
        run_ocr.engine = PaddleOCR(
            text_detection_model_name="PP-OCRv5_mobile_det",
            text_recognition_model_name="PP-OCRv5_mobile_rec",
            # Product photos are not documents. These three stages add a lot
            # of latency and are unnecessary for the upper-left hangtag.
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )
    def page_data(page):
        data = getattr(page, "json", None)
        if callable(data):
            data = data()
        if isinstance(data, str):
            import json
            data = json.loads(data)
        return data.get("res", data) if isinstance(data, dict) else {}

    def predict_text(path: str) -> tuple[str, dict[str, Any]]:
        result = run_ocr.engine.predict(path)
        text_parts: list[str] = []
        best: dict[str, Any] = {}
        for page in result:
            data = page_data(page)
            texts = data.get("rec_texts", [])
            text_parts.extend(str(x) for x in texts)
            if not best or len(texts) > len(best.get("rec_texts", [])):
                best = data
        return "\n".join(text_parts), best

    # First scan the complete image at reduced size. This makes the solution
    # independent of where the hangtag appears in the product photo.
    from PIL import Image
    scan_path = None
    crop_path = None
    try:
        with Image.open(image_path) as image:
            scale = min(1.0, 1000 / max(image.width, image.height))
            scan = image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                scan_path = tmp.name
            scan.save(scan_path, quality=88)
            scan_text, scan_data = predict_text(scan_path)

            # Locate the tag from printed field labels, then re-read that
            # region at original resolution so handwritten values are clearer.
            labels = ("货号", "款号", "编号", "型号", "颜色", "色号", "配色", "码数", "鞋码", "尺码", "号码")
            boxes = scan_data.get("rec_boxes", [])
            texts = scan_data.get("rec_texts", [])
            target_boxes = [box for text, box in zip(texts, boxes) if any(label in str(text) for label in labels) and len(box) >= 4]
            if not target_boxes:
                candidate_path = detect_tag_candidate(image_path)
                if candidate_path:
                    try:
                        candidate_text, _ = predict_text(candidate_path)
                        if candidate_text:
                            preview_path = None
                            if preview_dir:
                                Path(preview_dir).mkdir(parents=True, exist_ok=True)
                                preview_path = str(Path(preview_dir) / f"{Path(image_path).stem}.jpg")
                                from PIL import Image
                                with Image.open(candidate_path) as candidate_image:
                                    candidate_image.save(preview_path, quality=95)
                            return candidate_text, preview_path
                    finally:
                        if os.path.exists(candidate_path):
                            os.unlink(candidate_path)
                return scan_text, None

            xs = [float(v) for box in target_boxes for v in (box[0], box[2])]
            ys = [float(v) for box in target_boxes for v in (box[1], box[3])]
            left, right = min(xs) / scale, max(xs) / scale
            top, bottom = min(ys) / scale, max(ys) / scale
            tag_w, tag_h = right - left, bottom - top
            pad_x = max(100, int(max(tag_w, tag_h) * 0.65))
            pad_y = max(140, int(max(tag_w, tag_h) * 0.90))
            crop_box = (max(0, int(left - pad_x)), max(0, int(top - pad_y)), min(image.width, int(right + pad_x)), min(image.height, int(bottom + pad_y)))
            crop = image.crop(crop_box)
            # The provided tags are vertical; rotate the detected crop for
            # better handwritten recognition, without assuming its location.
            crop = crop.rotate(90, expand=True)
            preview_path = None
            if preview_dir:
                Path(preview_dir).mkdir(parents=True, exist_ok=True)
                preview_path = str(Path(preview_dir) / f"{Path(image_path).stem}.jpg")
                crop.save(preview_path, quality=95)
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
                crop_path = tmp.name
            crop.save(crop_path, quality=95)
        crop_text, _ = predict_text(crop_path)
        return crop_text or scan_text, preview_path
    finally:
        if scan_path and os.path.exists(scan_path):
            os.unlink(scan_path)
        if crop_path and os.path.exists(crop_path):
            os.unlink(crop_path)


def copy_row_style(ws, source_row: int, target_row: int, max_col: int) -> None:
    if source_row == target_row:
        return
    for col in range(1, max_col + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)
        if isinstance(target, MergedCell):
            continue
        if source.has_style:
            target._style = copy.copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy.copy(source.alignment)
        if source.protection:
            target.protection = copy.copy(source.protection)


def _merge_unique_values(values: list[str], *, size_values: bool = False) -> str:
    """Combine values in first-seen order, removing duplicates."""
    merged: list[str] = []
    for value in values:
        if size_values:
            value = clean_sizes(value)
            parts = re.findall(r"\d+(?:\.\d+)?", value)
        else:
            parts = [part.strip() for part in re.split(r"[，,、;/]+", str(value or ""))]
        for part in parts:
            if part and part not in merged:
                merged.append(part)
    return "，".join(merged)


def _split_color_values(value: str) -> list[str]:
    """Split merged color strings while keeping one color per output row."""
    return [clean_color(part) for part in re.split(r"[，,、;；\n\r]+", str(value or "")) if clean_color(part)]


def _color_key(value: str) -> str:
    return normalize_text(value).casefold()


def read_color_library(paths: list[str]) -> set[str]:
    """Read prior header-free color lists from the first column of each sheet."""
    known: set[str] = set()
    for path in paths:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            for sheet in workbook.worksheets:
                for (value,) in sheet.iter_rows(min_col=1, max_col=1, values_only=True):
                    for color in _split_color_values(str(value or "")):
                        known.add(_color_key(color))
        finally:
            workbook.close()
    return known


def export_color_list(output_path: str, rows: list[dict[str, Any]], library_paths: list[str]) -> int:
    """Create a header-free, one-column list of colors new to this export."""
    known = read_color_library(library_paths)
    colors: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for color in _split_color_values(str(row.get("color", ""))):
            key = _color_key(color)
            if key and key not in known and key not in seen:
                colors.append(color)
                seen.add(key)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "颜色清单"
    sheet.column_dimensions["A"].width = 20
    for index, color in enumerate(colors, start=1):
        sheet.cell(index, 1).value = color
    workbook.save(output_path)
    return len(colors)


def _export_gender_groups(size_value: str, gender: str = "") -> list[tuple[str, str]]:
    """Return export SKU suffixes and their corresponding gender size sets."""
    sizes = [int(value) for value in re.findall(r"\d+", clean_sizes(size_value))]
    size_set = set(sizes)
    women_sizes = set(range(35, 41))
    men_sizes = set(range(39, 45))
    # A full 35-44 range is a unisex source row.  Split it before consulting
    # the filename gender so the colors merge into the correct export rows.
    if women_sizes.issubset(size_set) and men_sizes.issubset(size_set):
        return [
            ("女", "，".join(str(size) for size in sorted(size_set & women_sizes))),
            ("男", "，".join(str(size) for size in sorted(size_set & men_sizes))),
        ]
    if gender in {"女", "男"}:
        return [(gender, clean_sizes(size_value))]
    if gender == "童":
        return [("童", clean_sizes(size_value))]
    groups: list[tuple[str, str]] = []
    if women_sizes.issubset(size_set):
        groups.append(("女", "，".join(str(size) for size in sorted(size_set & women_sizes))))
    if men_sizes.issubset(size_set):
        groups.append(("男", "，".join(str(size) for size in sorted(size_set & men_sizes))))
    return groups or [("", clean_sizes(size_value))]


def merge_rows_by_sku(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Merge export rows by SKU and gender range, leaving blank-SKU rows separate."""
    merged_rows: list[dict[str, Any]] = []
    by_sku: dict[str, dict[str, Any]] = {}
    for row in rows:
        sku = clean_sku(str(row.get("sku", "")))
        if not sku:
            merged_rows.append(copy.deepcopy(row))
            continue
        style_note = str(row.get("style_note", "")).strip()
        for suffix, gender_sizes in _export_gender_groups(str(row.get("size", "")), str(row.get("gender", ""))):
            export_sku = f"{sku}{suffix}{style_note}"
            key = export_sku.upper()
            if key not in by_sku:
                item = copy.deepcopy(row)
                item["sku"] = export_sku
                item["_export_colors"] = [str(row.get("color", ""))]
                item["_export_sizes"] = [gender_sizes]
                by_sku[key] = item
                merged_rows.append(item)
                continue
            item = by_sku[key]
            item["_export_colors"].append(str(row.get("color", "")))
            item["_export_sizes"].append(gender_sizes)
            for field in ("brand", "category", "price"):
                if not str(item.get(field, "")).strip() and str(row.get(field, "")).strip():
                    item[field] = row[field]
    for item in merged_rows:
        if "_export_colors" in item:
            item["color"] = _merge_unique_values(item.pop("_export_colors"))
            item["size"] = _merge_unique_values(item.pop("_export_sizes"), size_values=True)
    return merged_rows


def fill_template(template_path: str, output_path: str, rows: list[dict[str, Any]]) -> None:
    wb = load_workbook(template_path)
    ws = wb["商品数据"] if "商品数据" in wb.sheetnames else wb.worksheets[0]
    header_row = 1
    headers = {str(ws.cell(header_row, col).value).strip(): col for col in range(1, ws.max_column + 1) if ws.cell(header_row, col).value}

    # Remove template sample values but preserve the template's columns, widths and styles.
    for row in range(header_row + 1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None

    source_row = 2 if ws.max_row >= 2 else 1
    for index, item in enumerate(rows, start=2):
        if index > ws.max_row:
            ws.insert_rows(index)
        copy_row_style(ws, source_row, index, ws.max_column)
        brand = str(item.get("brand", "")).strip()
        sku = str(item.get("sku", "")).strip()
        spec1_name = str(item.get("spec1_name", "")).strip().rstrip("：:") or "颜色"
        values = {
            "商品名称": f"{brand}{sku}",
            "分类编号": item.get("category", ""),
            "商品品牌": item.get("brand", ""),
            "小单位": "双",
            "起订量": "1",
            "整倍订货": "1",
            "多规格1": f"{spec1_name}：{item.get('color', '')}" if item.get("color") else f"{spec1_name}：",
            "多规格2": f"码数：{item.get('size', '')}" if item.get("size") else "码数：",
            "小单位订货价": item.get("price", ""),
            "状态": "下架",
        }
        for name, value in values.items():
            col = headers.get(name)
            if col:
                ws.cell(index, col).value = value
    wb.save(output_path)


def recognize_one(path: str, defaults: dict[str, str], preview_dir: str | None = None) -> dict[str, Any]:
    last_row: dict[str, Any] = {}
    last_error = ""
    for attempt in range(1, 4):
        try:
            text, preview_path = run_ocr(path, preview_dir)
            parsed = parse_tag_text(text)
            parsed.update({
                "file": Path(path).name,
                "image_path": str(path),
                "brand": defaults.get("brand", ""),
                "category": defaults.get("category", ""),
                "price": defaults.get("price", ""),
                "preview_path": preview_path or "",
                "attempts": attempt,
            })
            errors = validate_row(parsed)
            core_errors = list(errors)
            parsed["quality_warnings"] = analyze_image_quality(path)
            parsed["warnings"] = []
            parsed["errors"] = core_errors + parsed["quality_warnings"]
            parsed["failure_reason"] = classify_failure(core_errors or parsed["quality_warnings"]) if (core_errors or parsed["quality_warnings"]) else ""
            parsed["recognition_status"] = "识别成功" if not core_errors else "识别失败"
            if not core_errors or attempt >= 3:
                return parsed
            last_error = "；".join(errors)
        except Exception as exc:
            last_error = str(exc)
    return {
        "file": Path(path).name,
        "image_path": str(path),
        "brand": defaults.get("brand", ""),
        "category": defaults.get("category", ""),
        "price": defaults.get("price", ""),
        "preview_path": "",
        "attempts": 3,
        "errors": [last_error or "OCR异常"],
        "quality_warnings": [],
        "warnings": [],
        "failure_reason": classify_failure([last_error or "OCR异常"], last_error),
        "recognition_status": "识别失败",
        "sku": "", "color": "", "size": "", "raw": "",
    }


def ocr_process_worker(files: list[str], defaults: dict[str, str], result_queue, preview_dir: str | None = None, start_index: int = 0, total_count: int | None = None) -> None:
    """Run PaddleOCR outside Tk's thread; PaddlePaddle can deadlock in a GUI thread on macOS."""
    try:
        total = total_count or len(files)
        for index, path in enumerate(files, start=start_index + 1):
            parsed = recognize_one(path, defaults, preview_dir)
            result_queue.put(("progress", index, len(files), parsed))
        result_queue.put(("done",))
    except Exception as exc:
        try:
            Path("/private/tmp/ai_product_assistant_error.log").write_text(traceback.format_exc(), encoding="utf-8")
        except Exception:
            pass
        result_queue.put(("error", str(exc)))


def ocr_worker_cli(manifest_path: str) -> None:
    """Worker mode used by the packaged app; results are streamed as JSONL."""
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    files = manifest["files"]
    defaults = manifest.get("defaults", {})
    preview_dir = manifest.get("preview_dir")
    start_index = int(manifest.get("start_index", 0))
    total_count = int(manifest.get("total_count", len(files)))
    try:
        for index, path in enumerate(files, start=start_index + 1):
            parsed = recognize_one(path, defaults, preview_dir)
            print(json.dumps({"kind": "progress", "index": index, "total": total_count, "row": parsed}, ensure_ascii=False), flush=True)
        print(json.dumps({"kind": "done"}, ensure_ascii=False), flush=True)
    except Exception as exc:
        print(json.dumps({"kind": "error", "message": str(exc)}, ensure_ascii=False), flush=True)


def preview_worker_cli(manifest_path: str) -> None:
    """Create a hangtag candidate preview for one image without filling the table."""
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    try:
        # Preview does not need to initialize PaddleOCR. Use the fast image
        # candidate detector first, then fall back to OCR only when needed.
        candidate_path = detect_tag_candidate(manifest["file"])
        if candidate_path:
            preview_dir = Path(manifest["preview_dir"])
            preview_dir.mkdir(parents=True, exist_ok=True)
            preview_path = preview_dir / Path(manifest["file"]).name
            shutil.copy2(candidate_path, preview_path)
            Path(candidate_path).unlink(missing_ok=True)
            print(json.dumps({"kind": "preview", "preview_path": str(preview_path), "method": "fast"}, ensure_ascii=False), flush=True)
            return
        _text, preview_path = run_ocr(manifest["file"], manifest["preview_dir"])
        print(json.dumps({"kind": "preview", "preview_path": preview_path or ""}, ensure_ascii=False), flush=True)
    except Exception as exc:
        print(json.dumps({"kind": "error", "message": str(exc)}, ensure_ascii=False), flush=True)


class App(AppBase):
    def __init__(self) -> None:
        super().__init__()
        self.title("AI 商品资料整理助手")
        self.geometry("1180x700")
        self.minsize(960, 560)
        self.template_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.brand_var = tk.StringVar()
        self.spec1_name_var = tk.StringVar()
        self.category_var = tk.StringVar()
        self.price_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择 Excel 模板和商品图片文件夹")
        self.rows: list[dict[str, Any]] = []
        self.selected_images: list[Path] = []
        self.ocr_process = None
        self.worker_manifest = None
        self.worker_queue = queue_module.Queue()
        self.worker_stderr = ""
        self.worker_done = False
        self.preview_dir = ""
        self.checkpoint_path = None
        self.status_label = None
        self.filter_var = tk.StringVar(value="全部")
        self.search_var = tk.StringVar()
        self.search_info_var = tk.StringVar()
        self.search_match_indices: list[int] = []
        self.color_library_info_var = tk.StringVar(value="未选择历史颜色表")
        self.color_library_paths: list[str] = []
        self.visible_indices: list[int] = []
        self.retry_mode = False
        self.retry_targets: list[Path] = []
        self.undo_stack: list[list[dict[str, Any]]] = []
        self.progress_var = tk.DoubleVar(value=0)
        self.stats_var = tk.StringVar(value="图片：0  已填码：0  待填码：0  格式错误：0")
        self._build_ui()
        self.search_var.trace_add("write", lambda *_args: self.refresh_tree())
        self.bind_all("<Control-z>", self.undo_last)
        self.bind_all("<Command-z>", self.undo_last)
        self.enable_drag_drop()
        self.after(300, self.offer_restore_autosave)

    def enable_drag_drop(self):
        if not DND_FILES:
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self.handle_drop)
            self.status_var.set("可拖入 Excel 模板、图片文件夹或单张图片")
        except Exception:
            pass

    def handle_drop(self, event):
        try:
            paths = [Path(p) for p in self.tk.splitlist(event.data)]
        except Exception:
            return
        if len(paths) == 1 and paths[0].is_file() and paths[0].suffix.lower() in {".xlsx", ".xlsm", ".xltx"}:
            self.template_var.set(str(paths[0])); return
        images = [p for p in paths if p.is_file() and p.suffix.lower() in IMAGE_EXTS]
        dirs = [p for p in paths if p.is_dir()]
        if len(images) == 1 and not dirs:
            self.folder_var.set(str(images[0].parent)); self.selected_images = images; return
        if len(dirs) == 1 and not images:
            self.folder_var.set(str(dirs[0])); self.selected_images = []
            return
        if images:
            self.folder_var.set(str(images[0].parent)); self.selected_images = images

    def _build_ui(self) -> None:
        top = ttk.LabelFrame(self, text="输入与默认值", padding=10)
        top.pack(fill="x", padx=12, pady=10)
        self._path_row(top, "Excel 模板", self.template_var, self.choose_template)
        self._path_row(top, "图片文件夹", self.folder_var, self.choose_folder, self.choose_single_image)
        form = ttk.Frame(top)
        form.pack(fill="x", pady=(8, 0))
        for label, var in [("商品品牌（可批量修改）", self.brand_var), ("多规格1名称（可批量修改）", self.spec1_name_var), ("分类编号（人工填写）", self.category_var), ("默认小单位订货价", self.price_var)]:
            ttk.Label(form, text=label).pack(side="left", padx=(0, 5))
            ttk.Entry(form, textvariable=var, width=18).pack(side="left", padx=(0, 14))

        actions = ttk.Frame(top)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="读取图片文件名", width=16, command=self.start_recognition).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="女码 35-40", width=14, command=lambda: self.fill_selected_size("女")).pack(side="left", padx=8)
        ttk.Button(actions, text="男码 39-44", width=14, command=lambda: self.fill_selected_size("男")).pack(side="left", padx=8)
        ttk.Button(actions, text="小童码 26-30", width=14, command=lambda: self.fill_selected_size("小童")).pack(side="left", padx=8)
        ttk.Button(actions, text="中童码 31-37", width=14, command=lambda: self.fill_selected_size("中童")).pack(side="left", padx=8)
        ttk.Button(actions, text="批量应用品牌/多规格1/分类/价格", width=30, command=self.apply_defaults).pack(side="left", padx=8)
        ttk.Button(actions, text="查看原图", width=12, command=self.show_original).pack(side="left", padx=8)
        ttk.Button(actions, text="导出表格", width=16, command=self.export).pack(side="right", padx=(8, 0))

        filter_bar = ttk.Frame(self)
        filter_bar.pack(fill="x", padx=12, pady=(0, 5))
        ttk.Label(filter_bar, text="结果筛选：").pack(side="left")
        filter_box = ttk.Combobox(filter_bar, textvariable=self.filter_var, values=["全部", "正常", "待填码", "格式错误"], state="readonly", width=12)
        filter_box.pack(side="left")
        filter_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_tree())
        ttk.Label(filter_bar, text="提示：女、男、男女自动填码；童鞋选中单行后点击小童或中童码。", foreground="#666666").pack(side="left", padx=15)
        ttk.Progressbar(filter_bar, variable=self.progress_var, maximum=100, length=180).pack(side="right", padx=(8, 0))
        ttk.Label(filter_bar, textvariable=self.stats_var).pack(side="right")

        search_bar = ttk.Frame(self)
        search_bar.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Label(search_bar, text="搜索定位：").pack(side="left")
        search_entry = ttk.Entry(search_bar, textvariable=self.search_var, width=34)
        search_entry.pack(side="left", padx=(5, 6))
        search_entry.bind("<Return>", self.focus_next_search_match)
        ttk.Button(search_bar, text="下一条", width=9, command=self.focus_next_search_match).pack(side="left")
        ttk.Button(search_bar, text="清除", width=8, command=self.clear_search).pack(side="left", padx=(6, 0))
        ttk.Label(search_bar, textvariable=self.search_info_var, foreground="#666666").pack(side="left", padx=12)

        color_library_bar = ttk.LabelFrame(self, text="颜色去重库（可选）", padding=(8, 4))
        color_library_bar.pack(fill="x", padx=12, pady=(0, 6))
        ttk.Button(color_library_bar, text="选择已有颜色表", width=16, command=self.choose_color_library_files).pack(side="left")
        ttk.Button(color_library_bar, text="清除选择", width=10, command=self.clear_color_library_files).pack(side="left", padx=(6, 0))
        ttk.Label(color_library_bar, textvariable=self.color_library_info_var, foreground="#666666").pack(side="left", padx=12)

        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        tree_body = ttk.Frame(table_frame)
        tree_body.pack(side="top", fill="both", expand=True)
        self.tree = ttk.Treeview(tree_body, columns=("file", "status", "reason", "sku", "color", "gender", "style_note", "size", "brand", "category", "price"), show="headings")
        headings = [("file", "图片"), ("status", "状态"), ("reason", "提示"), ("sku", "款号"), ("color", "颜色"), ("gender", "类别"), ("style_note", "款式说明"), ("size", "码数"), ("brand", "商品品牌"), ("category", "分类编号"), ("price", "小单位订货价")]
        for key, title in headings:
            self.tree.heading(key, text=title)
            widths = {"file": 150, "status": 90, "reason": 230, "sku": 140, "color": 130, "gender": 60, "style_note": 90, "size": 180, "brand": 110, "category": 110, "price": 120}
            self.tree.column(key, width=widths[key], anchor="w", stretch=False)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(tree_body, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        xscroll.pack(side="bottom", fill="x")
        self.tree.configure(xscrollcommand=xscroll.set)
        self.tree.tag_configure("failed", foreground="#d32f2f")
        self.tree.tag_configure("warning", foreground="#b26a00")
        self.tree.tag_configure("search_match", background="#1F6FEB", foreground="#FFFFFF")
        self.tree.bind("<ButtonRelease-1>", self.edit_cell)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.status_label = tk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padx=4, fg="#333333")
        self.status_label.pack(fill="x", side="bottom")

    def _path_row(self, parent, label, var, command, extra_command=None, preview_command=None, annotation_command=None):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label, width=12).pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row, text="选择", command=command).pack(side="right")
        if extra_command:
            ttk.Button(row, text="单张图片", command=extra_command).pack(side="right", padx=(0, 6))
        if preview_command:
            ttk.Button(row, text="预览吊牌区域", command=preview_command).pack(side="right", padx=(0, 6))
        if annotation_command:
            ttk.Button(row, text="手写标注", command=annotation_command).pack(side="right", padx=(0, 6))

    def choose_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm *.xltx"), ("所有文件", "*")])
        if path:
            self.template_var.set(path)

    def choose_color_library_files(self):
        paths = filedialog.askopenfilenames(
            title="选择已有颜色表（可多选）",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有文件", "*")],
        )
        if paths:
            self.color_library_paths = list(paths)
            self.color_library_info_var.set(f"已选择 {len(self.color_library_paths)} 份历史颜色表")

    def clear_color_library_files(self):
        self.color_library_paths = []
        self.color_library_info_var.set("未选择历史颜色表")

    def choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_var.set(path)
            self.selected_images = []
            self.restore_autosave(path)

    def choose_single_image(self):
        path = filedialog.askopenfilename(filetypes=[("图片", "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"), ("所有文件", "*")])
        if path:
            selected = Path(path)
            self.folder_var.set(str(selected.parent))
            self.selected_images = [selected]
            self.status_var.set(f"已选择单张图片：{selected.name}，可以读取文件名")

    def start_recognition(self, files_override=None):
        if not self.template_var.get() or not self.folder_var.get():
            messagebox.showwarning("缺少输入", "请先选择 Excel 模板和图片文件夹。")
            return
        files = files_override or self.selected_images or sorted(p for p in Path(self.folder_var.get()).iterdir() if p.suffix.lower() in IMAGE_EXTS)
        if not files:
            messagebox.showwarning("没有图片", "所选文件夹中没有支持的图片文件。")
            return
        files = [Path(p) for p in files]
        defaults = {"brand": self.brand_var.get(), "spec1_name": self.spec1_name_var.get(), "category": self.category_var.get(), "price": self.price_var.get()}
        self.rows = []
        for path in files:
            parsed = parse_image_filename(path)
            filename_error = parsed.pop("error")
            size = gender_size_value(parsed["gender"])
            row = {
                "file": path.name, "image_path": str(path), "sku": parsed["sku"], "color": parsed["color"],
                "gender": parsed["gender"], "style_note": parsed["style_note"], "size": size, "brand": defaults["brand"], "spec1_name": defaults["spec1_name"], "category": defaults["category"],
                "price": defaults["price"], "filename_error": filename_error, "errors": [], "warnings": [],
                "quality_warnings": [], "failure_reason": "", "recognition_status": "正常" if size and not filename_error else "待填码", "raw": "",
            }
            self.rows.append(row)
        self.progress_var.set(100)
        self.refresh_tree()
        invalid = sum(1 for row in self.rows if row.get("recognition_status") == "格式错误")
        filled = sum(1 for row in self.rows if row.get("recognition_status") == "正常")
        waiting = sum(1 for row in self.rows if row.get("recognition_status") == "待填码")
        self.status_var.set(f"已从文件名读取 {len(self.rows)} 张图片；自动填码 {filled} 张，待填码 {waiting} 张，格式错误 {invalid} 张。")
        if self.status_label:
            self.status_label.configure(fg="#d32f2f" if invalid else "#333333")

    def fill_selected_size(self, gender: str):
        selection = self.tree.selection()
        if len(selection) != 1:
            messagebox.showinfo("请选择单行", "请先在表格中选中一张图片，再填写码数。")
            return
        row = self.rows[self.visible_indices[self.tree.index(selection[0])]]
        if row.get("recognition_status") == "格式错误":
            messagebox.showwarning("文件名格式错误", "请先修改款号、颜色和性别后再填写码数。")
            return
        expected_categories = {"女": {"女"}, "男": {"男"}, "小童": {"童"}, "中童": {"童"}}
        if row.get("gender") and row["gender"] not in expected_categories[gender]:
            messagebox.showwarning("类别不匹配", f"该图片文件名标注为“{row['gender']}”，请点击对应的码段按钮。")
            return
        self.undo_stack.append(copy.deepcopy(self.rows))
        size_ranges = {
            "女": range(35, 41), "男": range(39, 45),
            "小童": range(26, 31), "中童": range(31, 38),
        }
        row["size"] = "，".join(str(size) for size in size_ranges[gender])
        if gender in {"女", "男"}:
            row["gender"] = gender
        row["recognition_status"] = "正常"
        self.autosave()
        self.refresh_tree()
        self.status_var.set(f"已填写 {row.get('file', '')}：{gender}码 {row['size']}")

    def read_worker_output(self):
        if not self.ocr_process:
            return
        try:
            if self.ocr_process.stdout:
                for line in self.ocr_process.stdout:
                    self.worker_queue.put(("line", line))
            if self.ocr_process.stderr:
                self.worker_stderr = self.ocr_process.stderr.read()
            self.ocr_process.wait()
            self.worker_queue.put(("exit", self.ocr_process.returncode))
        except Exception as exc:
            self.worker_queue.put(("reader_error", str(exc)))

    def poll_worker_queue(self):
        finished = False
        try:
            while True:
                kind, payload = self.worker_queue.get_nowait()
                if kind == "line":
                    try:
                        message = json.loads(payload)
                    except json.JSONDecodeError:
                        continue
                    message_kind = message.get("kind")
                    if message_kind == "progress":
                        new_row = message["row"]
                        if self.retry_mode:
                            match = next((i for i, row in enumerate(self.rows) if row.get("image_path") == new_row.get("image_path") or row.get("file") == new_row.get("file")), None)
                            if match is None:
                                self.rows.append(new_row)
                            else:
                                self.rows[match] = new_row
                        else:
                            self.rows.append(new_row)
                        self.refresh_tree()
                        self._save_checkpoint(message["index"])
                        self.progress_var.set((message["index"] / max(1, message["total"])) * 100)
                        self.status_var.set(f"正在识别 {message['index']}/{message['total']} 张图片……")
                    elif message_kind == "done":
                        finished = True
                        self.worker_done = True
                        self.retry_mode = False
                        self.retry_targets = []
                        if self.checkpoint_path:
                            self.checkpoint_path.unlink(missing_ok=True)
                        failed_count = sum(1 for row in self.rows if row.get("recognition_status") == "识别失败")
                        self.progress_var.set(100)
                        self.status_var.set(f"识别完成：{len(self.rows)} 张。失败 {failed_count} 张。双击表格单元格可人工校对。")
                        if self.status_label:
                            self.status_label.configure(fg="#d32f2f" if failed_count else "#333333")
                    elif message_kind == "error":
                        finished = True
                        self._mark_batch_failed(message.get("message", "OCR子程序异常"))
                elif kind == "reader_error":
                    finished = True
                    self._mark_batch_failed(payload)
                elif kind == "exit" and payload != 0 and not self.worker_done:
                    finished = True
                    detail = self.worker_stderr.strip()[-500:]
                    self._mark_batch_failed(detail or f"OCR 子程序退出码：{payload}")
        except queue_module.Empty:
            pass
        if not finished and self.ocr_process and self.ocr_process.poll() is None:
            self.after(100, self.poll_worker_queue)

    def _recognize_worker(self, files):
        output = []
        try:
            for index, path in enumerate(files, start=1):
                parsed = recognize_one(str(path), {"brand": self.brand_var.get(), "category": self.category_var.get(), "price": self.price_var.get()})
                output.append(parsed)
                self.after(0, lambda i=index, n=len(files): self.status_var.set(f"正在识别：{i}/{n}"))
        except Exception as exc:
            self.after(0, lambda: messagebox.showerror("识别失败", str(exc)))
            self.after(0, lambda: self.status_var.set("识别失败，请检查 PaddleOCR 安装或模型。"))
            return
        self.rows = output
        self.after(0, self.refresh_tree)
        self.after(0, lambda: self.status_var.set(f"识别完成：{len(output)} 张。双击表格单元格可人工校对。"))

    def row_matches_search(self, row: dict[str, Any], keyword: str) -> bool:
        if not keyword:
            return False
        searchable = " ".join(str(row.get(field, "")) for field in (
            "file", "sku", "color", "gender", "style_note", "size", "brand", "category", "price",
        ))
        normalized = normalize_text(searchable).casefold().replace("-", "").replace("_", "")
        return keyword in normalized

    def clear_search(self):
        self.search_var.set("")

    def focus_next_search_match(self, _event=None):
        if not self.search_match_indices:
            return "break"
        match_positions = [position for position, row_index in enumerate(self.visible_indices) if row_index in self.search_match_indices]
        if not match_positions:
            return "break"
        current_position = -1
        selection = self.tree.selection()
        if selection:
            current_position = self.tree.index(selection[0])
        next_position = next((position for position in match_positions if position > current_position), match_positions[0])
        item = self.tree.get_children()[next_position]
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.tree.see(item)
        return "break"

    def refresh_tree(self):
        try:
            scroll_position = self.tree.yview()[0]
        except Exception:
            scroll_position = 0.0
        self.apply_duplicate_checks()
        format_errors = sum(1 for row in self.rows if row.get("recognition_status") == "格式错误")
        waiting = sum(1 for row in self.rows if row.get("recognition_status") == "待填码")
        success = len(self.rows) - format_errors - waiting
        self.stats_var.set(f"图片：{len(self.rows)}  已填码：{success}  待填码：{waiting}  格式错误：{format_errors}")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.visible_indices = []
        self.search_match_indices = []
        selected_filter = self.filter_var.get()
        keyword = normalize_text(self.search_var.get()).casefold().replace("-", "").replace("_", "")
        for index, row in enumerate(self.rows):
            status = row.get("recognition_status", "正常")
            if status == "格式错误":
                status_text = "格式错误"
                tag = "failed"
            elif status == "待填码":
                status_text = "待填码"
                tag = "warning"
            else:
                status_text = "正常"
                tag = ""
            if selected_filter != "全部" and status_text != selected_filter:
                continue
            self.visible_indices.append(index)
            matched = self.row_matches_search(row, keyword)
            if matched:
                self.search_match_indices.append(index)
            detail = row.get("failure_reason", "")
            errors_text = "；".join(row.get("errors", []) + row.get("warnings", []))
            reason = f"{detail}：{errors_text}" if detail and errors_text else (detail or errors_text)
            tags = [tag] if tag else []
            if matched:
                tags.append("search_match")
            self.tree.insert("", "end", values=(row.get("file", ""), status_text, reason, row.get("sku", ""), row.get("color", ""), row.get("gender", ""), row.get("style_note", ""), f"码数：{row.get('size', '')}", row.get("brand", ""), row.get("category", ""), row.get("price", "")), tags=tuple(tags))
        if keyword:
            self.search_info_var.set(f"已高亮 {len(self.search_match_indices)} 条，未隐藏其他记录")
        else:
            self.search_info_var.set("")
        if self.tree.get_children():
            self.tree.yview_moveto(scroll_position)

    def apply_duplicate_checks(self):
        counts = {}
        for row in self.rows:
            sku = str(row.get("sku", "")).strip().upper()
            if sku:
                counts[sku] = counts.get(sku, 0) + 1
        for row in self.rows:
            errors = [error for error in row.get("errors", []) if error != "货号重复" and error != row.get("filename_error")]
            warnings = [warning for warning in row.get("warnings", []) if warning != "货号重复"]
            sku = str(row.get("sku", "")).strip().upper()
            if sku and counts.get(sku, 0) > 1:
                warnings.append("货号重复（提醒）")
            filename_error = row.get("filename_error", "")
            if filename_error:
                errors = [filename_error]
                status = "格式错误"
                failure_reason = "文件名格式异常"
            elif not str(row.get("size", "")).strip():
                errors = []
                status = "待填码"
                failure_reason = ""
            else:
                errors = validate_row(row)
                status = "正常" if not errors else "格式错误"
                failure_reason = classify_failure(errors) if errors else ""
            row["errors"] = errors
            row["warnings"] = warnings
            row["recognition_status"] = status
            row["failure_reason"] = failure_reason or ("重复款号提醒" if warnings else "")

    def apply_defaults(self):
        selected = self.tree.selection()
        target_indices = [self.visible_indices[self.tree.index(item)] for item in selected] if selected else list(range(len(self.rows)))
        for index in target_indices:
            row = self.rows[index]
            row["brand"] = self.brand_var.get()
            row["spec1_name"] = self.spec1_name_var.get()
            row["category"] = self.category_var.get()
            row["price"] = self.price_var.get()
        self.autosave()
        self.refresh_tree()

    def retry_failed(self):
        targets = [row for row in self.rows if row.get("recognition_status") == "识别失败"]
        if not targets:
            messagebox.showinfo("重试失败图片", "当前没有识别失败的图片。")
            return
        if self.ocr_process and self.ocr_process.poll() is None:
            messagebox.showinfo("正在识别", "当前批次还没有完成，请稍候。")
            return
        self.retry_mode = True
        self.retry_targets = [Path(row.get("image_path", "")) for row in targets if row.get("image_path")]
        self.start_recognition(self.retry_targets)

    def on_tree_select(self, _event=None):
        selection = self.tree.selection()
        if not selection:
            return
        index = self.visible_indices[self.tree.index(selection[0])]
        row = self.rows[index]
        errors = "；".join(row.get("errors", []) + row.get("warnings", []))
        self.status_var.set(f"{row.get('file', '')}：{errors or row.get('recognition_status', '正常')}")

    def show_original(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo("查看原图", "请先选择一行图片。")
            return
        index = self.visible_indices[self.tree.index(selection[0])]
        path = Path(self.rows[index].get("image_path", ""))
        if not path.exists():
            messagebox.showinfo("查看原图", "原图文件不存在。")
            return
        try:
            from PIL import Image, ImageTk
            window = tk.Toplevel(self)
            window.title(f"原图 - {path.name}")
            window.geometry("900x700")
            image = Image.open(path).convert("RGB")
            image.thumbnail((860, 620))
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(window, image=photo)
            label.image = photo
            label.pack(fill="both", expand=True, padx=10, pady=10)
        except Exception as exc:
            messagebox.showerror("查看原图失败", str(exc))

    def _save_checkpoint(self, completed: int):
        if self.retry_mode:
            self.autosave()
            return
        if not self.checkpoint_path:
            return
        try:
            files = [str(self.selected_images[0])] if self.selected_images else [str(p) for p in sorted(Path(self.folder_var.get()).iterdir()) if p.suffix.lower() in IMAGE_EXTS]
            self.checkpoint_path.write_text(json.dumps({"files": files, "completed": completed, "rows": self.rows}, ensure_ascii=False), encoding="utf-8")
            self.autosave()
        except Exception:
            pass

    def _mark_batch_failed(self, detail: str):
        self.status_var.set("识别失败，请检查红色标记的图片。")
        if self.status_label:
            self.status_label.configure(fg="#d32f2f")
        if detail:
            messagebox.showerror("识别失败", detail)

    def show_preview(self):
        selection = self.tree.selection()
        if not selection:
            if len(self.rows) == 1:
                index = 0
            else:
                messagebox.showinfo("吊牌预览", "请先完成识别并选择一行结果。")
                return
        else:
            index = self.visible_indices[self.tree.index(selection[0])]
        row = self.rows[index]
        preview = row.get("preview_path")
        if not preview or not Path(preview).exists():
            messagebox.showinfo("吊牌预览", "这张图片没有找到吊牌候选区域，请先重新识别。")
            return
        self.open_row_annotation(index, preview)

    def open_row_annotation(self, row_index: int, preview_path: str):
        row = self.rows[row_index]
        window = tk.Toplevel(self)
        window.title(f"吊牌校对 - {row.get('file', '')}")
        window.geometry("980x760")
        window.minsize(700, 520)
        window.transient(self)

        # The source image can be taller than the screen.  Put the image and
        # fields in one scrollable view so the editable fields and save button
        # always remain reachable.
        container = ttk.Frame(window)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        content = ttk.Frame(canvas)
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")

        def update_scroll_region(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def fit_content_width(event):
            canvas.itemconfigure(content_window, width=event.width)

        content.bind("<Configure>", update_scroll_region)
        canvas.bind("<Configure>", fit_content_width)

        def scroll_with_wheel(event):
            # Only consume wheel events belonging to this dialog; other app
            # windows retain their normal scrolling behaviour.
            try:
                if not window.winfo_exists() or event.widget.winfo_toplevel() != window:
                    return
            except tk.TclError:
                return
            if getattr(event, "delta", 0):
                canvas.yview_scroll(-1 * int(event.delta / 120 or (1 if event.delta > 0 else -1)), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-3, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(3, "units")
            return "break"

        canvas.bind_all("<MouseWheel>", scroll_with_wheel, add="+")
        canvas.bind_all("<Button-4>", scroll_with_wheel, add="+")
        canvas.bind_all("<Button-5>", scroll_with_wheel, add="+")

        image_frame = ttk.LabelFrame(content, text="吊牌候选区域（可直接校对下方文字）", padding=10)
        image_frame.pack(fill="both", expand=True, padx=12, pady=12)
        try:
            from PIL import Image, ImageTk
            image = Image.open(preview_path).convert("RGB")
            image.thumbnail((900, 560))
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(image_frame, image=photo)
            label.image = photo
            label.pack(fill="both", expand=True)
        except Exception as exc:
            ttk.Label(image_frame, text=str(exc)).pack(padx=20, pady=20)
        form = ttk.LabelFrame(content, text="手写内容校对", padding=10)
        form.pack(fill="x", padx=12, pady=(0, 12))
        sku_var = tk.StringVar(value=row.get("sku", ""))
        color_var = tk.StringVar(value=row.get("color", ""))
        size_var = tk.StringVar(value=row.get("size", ""))
        for label_text, variable in [("货号", sku_var), ("颜色", color_var), ("码数", size_var)]:
            line = ttk.Frame(form)
            line.pack(fill="x", pady=3)
            ttk.Label(line, text=label_text, width=8).pack(side="left")
            ttk.Entry(line, textvariable=variable).pack(side="left", fill="x", expand=True)
        buttons = ttk.Frame(form)
        buttons.pack(fill="x", pady=(8, 0))
        result_var = tk.StringVar(value="修改后点击保存")
        ttk.Label(buttons, textvariable=result_var, foreground="#666666").pack(side="left")
        def save_edit():
            self.undo_stack.append(copy.deepcopy(self.rows))
            row["sku"] = clean_sku(sku_var.get().strip())
            row["color"] = clean_color(color_var.get().strip())
            row["size"] = clean_sizes(size_var.get().strip())
            row["errors"] = validate_row(row)
            row["recognition_status"] = "识别成功" if not row["errors"] else "识别失败"
            row["failure_reason"] = classify_failure(row["errors"]) if row["errors"] else ""
            try:
                save_sample_record(row.get("image_path", ""), row["sku"], row["color"], row["size"])
            except Exception as exc:
                result_var.set(f"结果已修改，样本库保存失败：{exc}")
            else:
                result_var.set("已保存到当前结果和手写样本库")
            self.autosave()
            self.refresh_tree()
        ttk.Button(buttons, text="保存修改", command=save_edit).pack(side="right")
        ttk.Button(buttons, text="关闭", command=window.destroy).pack(side="right", padx=(0, 8))

    def start_single_image_preview(self, image_path: Path, callback=None):
        """Preview the candidate tag immediately after a single image is chosen."""
        preview_dir = tempfile.mkdtemp(prefix="ai_tag_preview_single_")
        manifest_file = tempfile.NamedTemporaryFile(prefix="ai_preview_", suffix=".json", delete=False)
        manifest_file.close()
        manifest_path = Path(manifest_file.name)
        manifest_path.write_text(json.dumps({"file": str(image_path), "preview_dir": preview_dir}, ensure_ascii=False), encoding="utf-8")
        bundled_python = Path("/Users/v/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3")
        source_candidates = [
            Path(__file__).resolve(),
            Path("/Users/v/Documents/Codex/2026-07-25/referenced-chatgpt-conversation-this-is-untrusted/outputs/ai_product_assistant/ai_product_assistant.py"),
        ]
        source_script = next((p for p in source_candidates if p.exists() and p.suffix == ".py"), source_candidates[-1])
        command = [str(bundled_python if bundled_python.exists() else Path(sys.executable)), "-u", str(source_script), "--preview-worker", str(manifest_path)]
        self.status_var.set("正在分析单张图片的吊牌区域……")
        if self.status_label:
            self.status_label.configure(fg="#333333")
        def task():
            try:
                # Run the fast candidate detector inside this background
                # thread first, avoiding a second Python process for preview.
                candidate_path = detect_tag_candidate(str(image_path))
                if candidate_path:
                    fast_preview = Path(preview_dir) / image_path.name
                    shutil.copy2(candidate_path, fast_preview)
                    Path(candidate_path).unlink(missing_ok=True)
                    message = {"kind": "preview", "preview_path": str(fast_preview), "method": "fast"}
                else:
                    completed = subprocess.run(command, capture_output=True, text=True, timeout=180)
                    message = json.loads(completed.stdout.strip().splitlines()[-1]) if completed.stdout.strip() else {"kind": "error", "message": completed.stderr[-500:]}
                self.after(0, lambda: self.finish_single_image_preview(message, image_path, callback))
            except Exception as exc:
                self.after(0, lambda: self.finish_single_image_preview({"kind": "error", "message": str(exc)}, image_path, callback))
            finally:
                manifest_path.unlink(missing_ok=True)
        threading.Thread(target=task, daemon=True).start()

    def finish_single_image_preview(self, message, image_path: Path, callback=None):
        if message.get("kind") != "preview" or not message.get("preview_path"):
            if callback:
                callback(None)
                return
            self.status_var.set("吊牌预览失败，请检查图片。")
            if self.status_label:
                self.status_label.configure(fg="#d32f2f")
            messagebox.showerror("吊牌预览失败", message.get("message", "没有找到吊牌候选区域"))
            return
        self.status_var.set(f"已生成吊牌候选区域：{image_path.name}")
        if callback:
            callback(message["preview_path"])
        else:
            self.show_preview_file(message["preview_path"], image_path.name)

    def show_preview_file(self, preview_path: str, title: str):
        window = tk.Toplevel(self)
        window.title(f"吊牌候选区域 - {title}")
        window.geometry("720x560")
        try:
            from PIL import Image, ImageTk
            image = Image.open(preview_path).convert("RGB")
            image.thumbnail((680, 480))
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(window, image=photo)
            label.image = photo
            label.pack(fill="both", expand=True, padx=10, pady=10)
        except Exception as exc:
            ttk.Label(window, text=str(exc)).pack(padx=20, pady=20)

    def open_annotation_tool(self):
        initial_image = None
        if len(self.selected_images) == 1:
            initial_image = self.selected_images[0]
        else:
            selection = self.tree.selection()
            if selection:
                index = self.visible_indices[self.tree.index(selection[0])]
                path = self.rows[index].get("image_path")
                if path:
                    initial_image = Path(path)
        AnnotationWindow(self, initial_image)

    def edit_cell(self, event):
        item_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item_id or column_id in {"#1", "#2", "#3"}:
            return
        row_index = self.visible_indices[self.tree.index(item_id)]
        field_map = {"#4": "sku", "#5": "color", "#6": "gender", "#7": "style_note", "#8": "size", "#9": "brand", "#10": "category", "#11": "price"}
        field = field_map.get(column_id)
        if not field:
            return
        self.open_cell_editor(row_index, field)

    def open_cell_editor(self, row_index, field):
        """Use a separate, reliable editor rather than an Entry overlaid on a Treeview."""
        labels = {
            "sku": "商品名称 / 吊牌货号",
            "color": "颜色",
            "gender": "类别（女、男、男女或童）",
            "style_note": "款式说明",
            "size": "码数",
            "brand": "商品品牌",
            "category": "分类编号",
            "price": "小单位订货价",
        }
        current = str(self.rows[row_index].get(field, "") or "")
        window = tk.Toplevel(self)
        window.title(f"修改{labels[field]}")
        window.resizable(False, False)
        window.transient(self)
        window.grab_set()
        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"{labels[field]}：").pack(anchor="w", pady=(0, 7))
        value = tk.StringVar(value=current)
        # Entry widgets do not support Tk's `undo` option.  Do not pass it here:
        # doing so aborts the click handler before the edit window can appear.
        editor = ttk.Entry(frame, textvariable=value, width=46)
        editor.pack(fill="x")
        hint = "颜色、码数不需要填写“颜色：”或“码数：”，程序会自动补上。" if field in {"color", "size"} else "修改后点击保存即可更新当前识别结果。"
        ttk.Label(frame, text=hint, foreground="#666666").pack(anchor="w", pady=(7, 12))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")

        def commit(_event=None):
            new_value = value.get().strip()
            new_value = new_value.removeprefix("颜色：").removeprefix("颜色:") if field == "color" else new_value
            new_value = new_value.removeprefix("码数：").removeprefix("码数:") if field == "size" else new_value
            self.undo_stack.append(copy.deepcopy(self.rows))
            self.rows[row_index][field] = new_value
            if field in {"sku", "color", "gender"}:
                self.rows[row_index]["filename_error"] = "" if self.rows[row_index].get("sku") and self.rows[row_index].get("color") and self.rows[row_index].get("gender") in {"女", "男", "男女", "童"} else "请填写款号、颜色和类别（女、男、男女或童）"
            self.autosave()
            self.refresh_tree()
            self.status_var.set(f"已修改：{labels[field]}")
            window.destroy()
            return "break"

        ttk.Button(buttons, text="取消", command=window.destroy, width=10).pack(side="right")
        ttk.Button(buttons, text="保存修改", command=commit, width=12).pack(side="right", padx=(0, 8))
        editor.bind("<Return>", commit)
        window.protocol("WM_DELETE_WINDOW", window.destroy)
        window.update_idletasks()
        self.update_idletasks()
        x = self.winfo_rootx() + max((self.winfo_width() - window.winfo_width()) // 2, 0)
        y = self.winfo_rooty() + max((self.winfo_height() - window.winfo_height()) // 2, 0)
        window.geometry(f"+{x}+{y}")
        editor.focus_set()
        editor.selection_range(0, tk.END)

    def undo_last(self, _event=None):
        if not self.undo_stack:
            return "break"
        self.rows = self.undo_stack.pop()
        self.refresh_tree()
        self.autosave()
        self.status_var.set("已撤销上一次人工修改")
        return "break"

    def export(self):
        if not self.rows:
            messagebox.showwarning("没有数据", "请先读取图片文件名并填写码数。")
            return
        choice = self.choose_export_scope()
        if not choice:
            return
        if choice == "全部":
            export_rows = self.rows
        elif choice == "已填码":
            export_rows = [row for row in self.rows if row.get("recognition_status") == "正常"]
        else:
            export_rows = [row for row in self.rows if row.get("recognition_status") != "正常"]
        if not export_rows:
            messagebox.showinfo("没有可导出的数据", f"当前没有符合“{choice}”条件的数据。")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"商品资料识别结果_{choice}.xlsx")
        if not path:
            return
        try:
            image_row_count = len(export_rows)
            export_rows = merge_rows_by_sku(export_rows)
            fill_template(self.template_var.get(), path, export_rows)
            color_path = str(Path(path).with_name(f"{Path(path).stem}_颜色清单.xlsx"))
            color_count = export_color_list(color_path, export_rows, self.color_library_paths)
            messagebox.showinfo(
                "导出完成",
                f"已按款号合并：{image_row_count} 张图片 → {len(export_rows)} 个导出项。\n"
                f"颜色清单：{color_count} 个唯一颜色。\n\n已生成：\n{path}\n{color_path}",
            )
            self.status_var.set(f"导出完成：{len(export_rows)} 个款号项；颜色清单 {color_count} 个唯一颜色")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def choose_export_scope(self):
        window = tk.Toplevel(self)
        window.title("选择导出范围")
        window.geometry("300x170")
        value = tk.StringVar(value="全部")
        ttk.Label(window, text="请选择导出内容：").pack(pady=(18, 8))
        ttk.Combobox(window, textvariable=value, values=["全部", "已填码", "待处理"], state="readonly", width=16).pack()
        result = {"value": None}
        def confirm():
            result["value"] = value.get()
            window.destroy()
        ttk.Button(window, text="确定", command=confirm).pack(pady=15)
        window.grab_set()
        self.wait_window(window)
        return result["value"]

    def task_state(self):
        return {"template": self.template_var.get(), "folder": self.folder_var.get(), "selected_images": [str(p) for p in self.selected_images], "brand": self.brand_var.get(), "spec1_name": self.spec1_name_var.get(), "category": self.category_var.get(), "price": self.price_var.get(), "color_library_paths": self.color_library_paths, "rows": self.rows}

    def save_task_to(self, path: str):
        Path(path).write_text(json.dumps(self.task_state(), ensure_ascii=False, indent=2), encoding="utf-8")

    def save_task(self):
        path = filedialog.asksaveasfilename(defaultextension=".aitask", filetypes=[("商品整理任务", "*.aitask")], initialfile="商品资料任务.aitask")
        if path:
            try:
                self.save_task_to(path)
                self.status_var.set(f"任务已保存：{path}")
            except Exception as exc:
                messagebox.showerror("保存任务失败", str(exc))

    def open_task(self):
        path = filedialog.askopenfilename(filetypes=[("商品整理任务", "*.aitask"), ("所有文件", "*")])
        if not path:
            return
        try:
            state = json.loads(Path(path).read_text(encoding="utf-8"))
            self.template_var.set(state.get("template", "")); self.folder_var.set(state.get("folder", ""))
            self.selected_images = [Path(p) for p in state.get("selected_images", [])]
            self.brand_var.set(state.get("brand", "")); self.spec1_name_var.set(state.get("spec1_name", "")); self.category_var.set(state.get("category", "")); self.price_var.set(state.get("price", ""))
            self.color_library_paths = [str(p) for p in state.get("color_library_paths", []) if Path(p).exists()]
            self.color_library_info_var.set(f"已选择 {len(self.color_library_paths)} 份历史颜色表" if self.color_library_paths else "未选择历史颜色表")
            self.rows = state.get("rows", [])
            self.undo_stack = []
            self.refresh_tree(); self.status_var.set(f"任务已打开：{Path(path).name}")
        except Exception as exc:
            messagebox.showerror("打开任务失败", str(exc))

    def new_task(self):
        has_content = bool(self.rows or self.template_var.get() or self.folder_var.get() or self.brand_var.get() or self.spec1_name_var.get() or self.category_var.get() or self.price_var.get() or self.color_library_paths)
        if has_content and not messagebox.askyesno("新建任务", "当前任务内容将被清空，确定新建吗？"):
            return
        self.template_var.set("")
        self.folder_var.set("")
        self.selected_images = []
        self.rows = []
        self.undo_stack = []
        self.brand_var.set("")
        self.spec1_name_var.set("")
        self.category_var.set("")
        self.price_var.set("")
        self.clear_color_library_files()
        self.filter_var.set("全部")
        self.visible_indices = []
        self.checkpoint_path = None
        self.retry_mode = False
        self.retry_targets = []
        self.refresh_tree()
        self.status_var.set("已新建空白任务，请选择 Excel 模板和图片")
        if self.status_label:
            self.status_label.configure(fg="#333333")

    def autosave(self):
        if not self.folder_var.get():
            return
        try:
            Path(self.folder_var.get(), ".ai_product_autosave.json").write_text(json.dumps(self.task_state(), ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    def offer_restore_autosave(self):
        folder = self.folder_var.get()
        if folder:
            return
        # Autosave is discovered when a task folder is selected; this hook keeps startup quiet.

    def restore_autosave(self, folder: str):
        path = Path(folder) / ".ai_product_autosave.json"
        if not path.exists() or self.rows:
            return
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
            saved_rows = state.get("rows", [])
            if saved_rows and messagebox.askyesno("恢复自动保存", f"发现该文件夹有 {len(saved_rows)} 条自动保存结果，是否恢复？"):
                self.template_var.set(state.get("template", self.template_var.get()))
                self.selected_images = [Path(p) for p in state.get("selected_images", [])]
                self.brand_var.set(state.get("brand", "")); self.spec1_name_var.set(state.get("spec1_name", "")); self.category_var.set(state.get("category", "")); self.price_var.set(state.get("price", ""))
                self.color_library_paths = [str(p) for p in state.get("color_library_paths", []) if Path(p).exists()]
                self.color_library_info_var.set(f"已选择 {len(self.color_library_paths)} 份历史颜色表" if self.color_library_paths else "未选择历史颜色表")
                self.rows = saved_rows
                self.refresh_tree()
        except Exception:
            pass


class AnnotationWindow(tk.Toplevel):
    """Manual label collection tool for building a handwriting fine-tuning set."""
    def __init__(self, parent, initial_image: Path | None = None):
        super().__init__(parent)
        self.parent_app = parent
        self.title("手写样本标注")
        self.geometry("980x650")
        self.folder = ""
        self.files: list[Path] = []
        self.annotations: dict[str, dict[str, str]] = {}
        self.index = -1
        self.preview_image = None
        self.folder_var = tk.StringVar()
        self.sku_var = tk.StringVar()
        self.color_var = tk.StringVar()
        self.size_var = tk.StringVar()
        self.count_var = tk.StringVar(value="尚未选择图片文件夹")
        SAMPLE_LIBRARY_DIR.mkdir(parents=True, exist_ok=True)
        SAMPLE_LIBRARY_IMAGES.mkdir(parents=True, exist_ok=True)
        self.build_ui()
        if initial_image and initial_image.exists():
            self.load_files(str(initial_image.parent), [initial_image])

    def build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="图片文件夹").pack(side="left")
        ttk.Entry(top, textvariable=self.folder_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(top, text="选择文件夹", command=self.choose_folder).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="选择单张图片", command=self.choose_image).pack(side="left")
        ttk.Button(top, text="打开样本库", command=self.open_library).pack(side="left", padx=(6, 0))
        ttk.Label(top, textvariable=self.count_var).pack(side="right", padx=(10, 0))

        body = ttk.Frame(self, padding=(10, 0, 10, 10))
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body)
        left.pack(side="left", fill="y", padx=(0, 10))
        self.listbox = tk.Listbox(left, width=32, exportselection=False)
        self.listbox.pack(side="left", fill="y", expand=True)
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.listbox.yview)
        scroll.pack(side="right", fill="y")
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.bind("<<ListboxSelect>>", self.select_image)

        right = ttk.Frame(body)
        right.pack(side="left", fill="both", expand=True)
        self.image_label = ttk.Label(right, text="选择图片后预览", anchor="center")
        self.image_label.pack(fill="both", expand=True)
        form = ttk.LabelFrame(right, text="吊牌文字标注", padding=10)
        form.pack(fill="x", pady=(10, 0))
        for label, var in [("货号", self.sku_var), ("颜色", self.color_var), ("码数", self.size_var)]:
            row = ttk.Frame(form)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label, width=8).pack(side="left")
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        buttons = ttk.Frame(form)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="保存当前标注", command=self.save_current).pack(side="left")
        ttk.Button(buttons, text="上一张", command=self.previous).pack(side="left", padx=8)
        ttk.Button(buttons, text="下一张", command=self.next).pack(side="left")
        ttk.Label(form, text=f"样本库：{SAMPLE_LIBRARY_DIR}（会集中保存标注记录和样本图片）", foreground="#666666").pack(anchor="w", pady=(8, 0))

    def choose_folder(self):
        folder = filedialog.askdirectory(parent=self)
        if not folder:
            return
        self.load_files(folder, sorted(p for p in Path(folder).iterdir() if p.suffix.lower() in IMAGE_EXTS))

    def choose_image(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("图片", "*.jpg *.jpeg *.png *.webp *.bmp *.tif *.tiff"), ("所有文件", "*")])
        if not path:
            return
        self.load_files(str(Path(path).parent), [Path(path)])

    def load_files(self, folder, files):
        self.folder = folder
        self.folder_var.set(folder)
        self.files = files
        self.annotations = {}
        output = Path(folder) / "handwriting_samples.jsonl"
        for source in (SAMPLE_LIBRARY_FILE, output):
            if not source.exists():
                continue
            try:
                for line in source.read_text(encoding="utf-8").splitlines():
                    item = json.loads(line)
                    key = item.get("source_image", item.get("image", ""))
                    if key:
                        self.annotations[key] = item
            except Exception:
                continue
        self.listbox.delete(0, tk.END)
        for path in self.files:
            self.listbox.insert(tk.END, path.name)
        self.count_var.set(f"共 {len(self.files)} 张，已标注 {len(self.annotations)} 张")
        if self.files:
            self.listbox.selection_set(0)
            self.listbox.event_generate("<<ListboxSelect>>")

    def select_image(self, _event=None):
        selection = self.listbox.curselection()
        if not selection:
            return
        self.index = selection[0]
        path = self.files[self.index]
        try:
            from PIL import Image, ImageTk
            image = Image.open(path).convert("RGB")
            image.thumbnail((620, 410))
            self.preview_image = ImageTk.PhotoImage(image)
            self.image_label.configure(image=self.preview_image, text="")
        except Exception as exc:
            self.image_label.configure(image="", text=str(exc))
        item = self.annotations.get(str(path), {})
        self.sku_var.set(item.get("sku", ""))
        self.color_var.set(item.get("color", ""))
        self.size_var.set(item.get("size", ""))

    def save_current(self):
        if self.index < 0 or not self.files:
            return
        path = self.files[self.index]
        item = {"image": str(path), "sku": self.sku_var.get().strip(), "color": self.color_var.get().strip(), "size": clean_sizes(self.size_var.get().strip())}
        self.annotations[str(path)] = item
        output = Path(self.folder) / "handwriting_samples.jsonl"
        output.write_text("\n".join(json.dumps(self.annotations[key], ensure_ascii=False) for key in sorted(self.annotations)) + "\n", encoding="utf-8")
        try:
            library_name = hashlib.sha1(str(path.resolve()).encode("utf-8")).hexdigest()[:12] + "_" + path.name
            library_image = SAMPLE_LIBRARY_IMAGES / library_name
            if not library_image.exists():
                shutil.copy2(path, library_image)
            library_items = {}
            if SAMPLE_LIBRARY_FILE.exists():
                for line in SAMPLE_LIBRARY_FILE.read_text(encoding="utf-8").splitlines():
                    old = json.loads(line)
                    key = old.get("source_image", old.get("image", ""))
                    if key:
                        library_items[key] = old
            library_item = dict(item)
            library_item["image"] = str(library_image)
            library_item["source_image"] = str(path)
            library_items[str(path)] = library_item
            SAMPLE_LIBRARY_FILE.write_text("\n".join(json.dumps(library_items[key], ensure_ascii=False) for key in sorted(library_items)) + "\n", encoding="utf-8")
        except Exception as exc:
            messagebox.showwarning("样本库保存提醒", f"本地标注已保存，但集中样本库保存失败：\n{exc}")
        self.count_var.set(f"共 {len(self.files)} 张，已标注 {len(self.annotations)} 张")

    def preview_tag_position(self):
        if self.index < 0 or not self.files:
            messagebox.showinfo("预览吊牌位置", "请先选择一张图片。", parent=self)
            return
        image_path = self.files[self.index]
        self.count_var.set(f"正在分析吊牌位置：{image_path.name}，请稍候……")
        if hasattr(self.parent_app, "start_single_image_preview"):
            self.parent_app.start_single_image_preview(image_path, self.show_tag_preview)

    def show_tag_preview(self, preview_path: str):
        if not preview_path:
            self.count_var.set("吊牌预览失败，请检查图片后重试")
            self.image_label.configure(image="", text="未找到吊牌候选区域")
            return
        try:
            from PIL import Image, ImageTk
            image = Image.open(preview_path).convert("RGB")
            image.thumbnail((620, 540))
            self.preview_image = ImageTk.PhotoImage(image)
            self.image_label.configure(image=self.preview_image, text="吊牌候选区域（再次点击左侧图片可恢复原图）")
            self.count_var.set(f"已显示吊牌候选区域：{self.files[self.index].name}")
        except Exception as exc:
            self.count_var.set("吊牌预览失败")
            self.image_label.configure(image="", text=f"吊牌预览失败：{exc}")

    def open_library(self):
        SAMPLE_LIBRARY_DIR.mkdir(parents=True, exist_ok=True)
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(SAMPLE_LIBRARY_DIR)])
        elif os.name == "nt":
            os.startfile(str(SAMPLE_LIBRARY_DIR))
        else:
            subprocess.Popen(["xdg-open", str(SAMPLE_LIBRARY_DIR)])

    def previous(self):
        if self.index > 0:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.index - 1)
            self.listbox.event_generate("<<ListboxSelect>>")

    def next(self):
        self.save_current()
        if self.index < len(self.files) - 1:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(self.index + 1)
            self.listbox.event_generate("<<ListboxSelect>>")


if __name__ == "__main__":
    if "--ocr-worker" in sys.argv:
        ocr_worker_cli(sys.argv[sys.argv.index("--ocr-worker") + 1])
    elif "--preview-worker" in sys.argv:
        preview_worker_cli(sys.argv[sys.argv.index("--preview-worker") + 1])
    else:
        multiprocessing.freeze_support()
        App().mainloop()
