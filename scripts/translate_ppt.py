#!/usr/bin/env python3
import argparse
import os
import sys
import re
from io import BytesIO
from typing import Tuple

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from langdetect import detect
    from PIL import Image
    import pytesseract
except Exception as e:
    print(f"[ERROR] Missing dependencies: {e}")
    sys.exit(1)

# Argos Translate (offline)
ARGOS_AVAILABLE = False
try:
    import argostranslate.package as argos_package
    import argostranslate.translate as argos_translate
    ARGOS_AVAILABLE = True
except Exception:
    ARGOS_AVAILABLE = False

CHINESE_CHAR_RANGE = re.compile(r"[\u4e00-\u9fff]")

def is_chinese_text(text: str) -> bool:
    if not text:
        return False
    return bool(CHINESE_CHAR_RANGE.search(text))

def ensure_argos_model(from_code: str, to_code: str = "zh") -> bool:
    if not ARGOS_AVAILABLE:
        return False
    try:
        argos_package.update()
        installed = []
        for lang in argos_translate.get_installed_languages():
            for to_lang in lang.translations:
                installed.append((lang.code, to_lang.code))
        if (from_code, to_code) in installed:
            return True
        available = argos_package.get_available_packages()
        candidates = [p for p in available if p.from_code == from_code and p.to_code == to_code]
        if not candidates:
            print(f"[WARN] No Argos package available for {from_code}->{to_code}")
            return False
        pkg = candidates[0]
        path = pkg.download()
        argos_package.install_from_path(path)
        print(f"[INFO] Installed Argos model {from_code}->{to_code}")
        return True
    except Exception as e:
        print(f"[WARN] Argos model install failed: {e}")
        return False

def argos_translate_text(text: str, force: bool) -> str:
    if not ARGOS_AVAILABLE:
        return text
    if not text or text.strip() == "":
        return text
    if not force and is_chinese_text(text):
        return text
    src_lang = "en"
    try:
        d = detect(text)
        src_lang = (d or "en").split("-")[0]
    except Exception:
        src_lang = "en"
    # Ensure model; try src->zh then fallback en->zh
    if not ensure_argos_model(src_lang, "zh"):
        if src_lang != "en" and ensure_argos_model("en", "zh"):
            src_lang = "en"
        else:
            print(f"[WARN] No Argos model for {src_lang}->zh; leaving text unchanged.")
            return text
    try:
        out = argos_translate.translate(text, from_code=src_lang, to_code="zh")
        return out
    except Exception as e:
        print(f"[WARN] Argos translation failed: {e}")
        return text

def translate_text(text: str, force: bool) -> str:
    return argos_translate_text(text, force=force)

def ocr_image_to_text(image_blob: bytes) -> str:
    try:
        img = Image.open(BytesIO(image_blob))
        txt = pytesseract.image_to_string(img)
        return txt
    except Exception as e:
        print(f"[WARN] OCR failed: {e}")
        return ""

def validate_paths(src: str, out: str) -> Tuple[str, str]:
    src = (src or "").strip()
    out = (out or "").strip()
    if not src:
        print("[ERROR] --source path is empty after trimming spaces.")
        sys.exit(2)
    if not os.path.exists(src):
        print(f"[ERROR] Source file not found: '{src}'. Please check filename and path.")
        # Show directory listing to help debugging
        try:
            print("[INFO] Current directory listing:")
            for p in os.listdir("."):
                print(" -", p)
        except Exception:
            pass
        sys.exit(3)
    return src, out

def process_presentation(src_path: str, out_path: str, enable_ocr: bool = False, force: bool = False):
    src_path, out_path = validate_paths(src_path, out_path)
    print(f"[INFO] Start translation")
    print(f"[INFO] Source: {src_path}")
    print(f"[INFO] Output: {out_path}")
    print(f"[INFO] OCR enabled: {enable_ocr}")
    print(f"[INFO] Force translate: {force}")

    prs = Presentation(src_path)
    slide_count = len(prs.slides)
    changed_text_items = 0
    processed_shapes = 0

    for si, slide in enumerate(prs.slides, start=1):
        slide_changes = 0
        for shape in slide.shapes:
            processed_shapes += 1
            # Text frames
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                original = shape.text
                translated = translate_text(original, force=force)
                if translated != original:
                    shape.text = translated
                    changed_text_items += 1
                    slide_changes += 1
            # Tables
            if getattr(shape, "has_table", False) and shape.has_table:
                tbl = shape.table
                for r in range(len(tbl.rows)):
                    for c in range(len(tbl.columns)):
                        cell = tbl.cell(r, c)
                        original = cell.text
                        translated = translate_text(original, force=force)
                        if translated != original:
                            cell.text = translated
                            changed_text_items += 1
                            slide_changes += 1
            # Pictures (OCR optional)
            if enable_ocr and getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = shape.image.blob
                    ocr_txt = ocr_image_to_text(blob)
                    if ocr_txt and (force or not is_chinese_text(ocr_txt)):
                        translated = translate_text(ocr_txt, force=force)
                        alt = (getattr(shape, "alternative_text", "") or "").strip()
                        note = f"OCR: {ocr_txt.strip()}\nTranslation: {translated.strip()}"
                        try:
                            shape.alternative_text = f"{alt}\n{note}" if alt else note
                        except Exception:
                            pass
                except Exception as e:
                    print(f"[WARN] OCR processing failed: {e}")
        # Speaker notes
        try:
            if getattr(slide, "has_notes_slide", False) and slide.has_notes_slide:
                notes_slide = slide.notes_slide
                if getattr(notes_slide, "notes_text_frame", None):
                    original = notes_slide.notes_text_frame.text
                    translated = translate_text(original, force=force)
                    if translated != original:
                        notes_slide.notes_text_frame.text = translated
                        changed_text_items += 1
                        slide_changes += 1
        except Exception as e:
            print(f"[WARN] Notes processing failed on slide {si}: {e}")

        print(f"[INFO] Slide {si}/{slide_count}: changes={slide_changes}")

    prs.save(out_path)
    print(f"[INFO] Saved translated PPT: {out_path}")
    print(f"[INFO] Slides: {slide_count}, Shapes processed: {processed_shapes}, Text items changed: {changed_text_items}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Translate non-Chinese PPT text to Simplified Chinese (offline)")
    parser.add_argument("--source", default="01计算机设计导论.pptx", help="Source PPTX path")
    parser.add_argument("--output", default="01计算机设计导论_zh-CN.pptx", help="Output PPTX path")
    parser.add_argument("--enable-ocr", action="store_true", help="Enable OCR for pictures")
    parser.add_argument("--force", action="store_true", help="Force translation even if detection says Chinese/mixed")
    args = parser.parse_args()

    process_presentation(args.source, args.output, enable_ocr=args.enable_ocr, force=args.force)
