#!/usr/bin/env python3
import argparse
import os
import sys
import re
from io import BytesIO

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from langdetect import detect
    from PIL import Image
    import pytesseract
except Exception as e:
    print(f"[ERROR] Missing dependencies: {e}")
    sys.exit(1)

# Optional: offline translator Argos Translate
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
        argos_package.update_package_index()
        installed = []
        for lang in argos_translate.get_installed_languages():
            for to_lang in lang.translations_from(lang):
                installed.append((lang.code, to_lang.code))
        if (from_code, to_code) in installed:
            return True
        available = argos_package.get_available_packages()
        candidates = [p for p in available if p.from_code == from_code and p.to_code == to_code]
        if not candidates:
            return False
        pkg = candidates[0]
        path = argos_package.download_and_install_package(pkg)
        return True
    except Exception as e:
        print(f"[WARN] Argos model install failed: {e}")
        return False


def argos_translate_text(text: str) -> str:
    if not ARGOS_AVAILABLE:
        return text
    if not text or text.strip() == "":
        return text
    if is_chinese_text(text):
        return text
    src_lang = "en"
    try:
        d = detect(text)
        src_lang = (d or "en").split("-")[0]
    except Exception:
        src_lang = "en"
    # Ensure model
    if not ensure_argos_model(src_lang, "zh"):
        if src_lang != "en" and ensure_argos_model("en", "zh"):
            src_lang = "en"
        else:
            print(f"[WARN] No Argos model for {src_lang}->zh; leaving text unchanged.")
            return text
    try:
        return argos_translate.translate(text, src_lang, "zh")
    except Exception as e:
        print(f"[WARN] Argos translation failed: {e}")
        return text


def translate_text(text: str) -> str:
    # Offline-only: Argos Translate
    return argos_translate_text(text)


def ocr_image_to_text(image_blob: bytes) -> str:
    try:
        img = Image.open(BytesIO(image_blob))
        txt = pytesseract.image_to_string(img)
        return txt
    except Exception as e:
        print(f"[WARN] OCR failed: {e}")
        return ""


def process_presentation(src_path: str, out_path: str, enable_ocr: bool = False):
    prs = Presentation(src_path)

    for si, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            # Text frames
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                original = shape.text
                translated = translate_text(original)
                if translated != original:
                    shape.text = translated
            # Tables
            if hasattr(shape, "has_table") and shape.has_table:
                tbl = shape.table
                for r in range(len(tbl.rows)):
                    for c in range(len(tbl.columns)):
                        cell = tbl.cell(r, c)
                        original = cell.text
                        translated = translate_text(original)
                        if translated != original:
                            cell.text = translated
            # Pictures (OCR optional)
            if enable_ocr and shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = shape.image.blob
                    ocr_txt = ocr_image_to_text(blob)
                    if ocr_txt and not is_chinese_text(ocr_txt):
                        translated = translate_text(ocr_txt)
                        alt = (shape.alternative_text or "").strip()
                        note = f"OCR: {ocr_txt.strip()}\nTranslation: {translated.strip()}"
                        shape.alternative_text = f"{alt}\n{note}" if alt else note
                except Exception as e:
                    print(f"[WARN] OCR processing failed: {e}")
        # Speaker notes
        try:
            if hasattr(slide, "has_notes_slide") and slide.has_notes_slide:
                notes_slide = slide.notes_slide
                if hasattr(notes_slide, "notes_text_frame") and notes_slide.notes_text_frame:
                    original = notes_slide.notes_text_frame.text
                    translated = translate_text(original)
                    if translated != original:
                        notes_slide.notes_text_frame.text = translated
        except Exception as e:
            print(f"[WARN] Notes processing failed on slide {si}: {e}")

    prs.save(out_path)
    print(f"[INFO] Saved translated PPT: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Translate non-Chinese PPT text to Simplified Chinese (offline)")
    parser.add_argument("--source", default="01计算机设计导论.pptx", help="Source PPTX path")
    parser.add_argument("--output", default="01计算机设计导论_zh-CN.pptx", help="Output PPTX path")
    parser.add_argument("--enable-ocr", action="store_true", help="Enable OCR for pictures")
    args = parser.parse_args()

    process_presentation(args.source, args.output, enable_ocr=args.enable_ocr)
