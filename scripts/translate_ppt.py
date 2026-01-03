#!/usr/bin/env python3
import argparse
import os
import sys
import re
from io import BytesIO

try:
    import deepl
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from langdetect import detect, LangDetectException
    from PIL import Image
    import pytesseract
except (ImportError, ModuleNotFoundError) as e:
    print(f"[ERROR] Missing dependencies: {e}")
    sys.exit(1)

CHINESE_CHAR_RANGE = re.compile(r"[\u4e00-\u9fff]")


def is_chinese_text(text: str) -> bool:
    if not text:
        return False
    # Consider text Chinese if it contains at least one CJK Unified Ideograph
    return bool(CHINESE_CHAR_RANGE.search(text))


def translate_text(translator: deepl.Translator, text: str) -> str:
    if not text or text.strip() == "":
        return text
    try:
        # Skip translation if text appears to contain Chinese
        if is_chinese_text(text):
            return text
        # If language detection yields Chinese, skip
        try:
            lang = detect(text)
            if lang.lower().startswith("zh"):
                return text
        except LangDetectException:
            # If detection fails, attempt translation anyway
            pass
        result = translator.translate_text(text, target_lang="ZH")
        return str(result)
    except Exception as e:
        print(f"[WARN] Translation failed for text: {text[:60]}... Error: {e}")
        return text


def ocr_image_to_text(image_blob: bytes) -> str:
    try:
        img = Image.open(BytesIO(image_blob))
        txt = pytesseract.image_to_string(img)
        return txt
    except Exception as e:
        print(f"[WARN] OCR failed: {e}")
        return ""


def process_presentation(src_path: str, out_path: str, enable_ocr: bool = False):
    deepl_key = os.getenv("DEEPL_API_KEY")
    if not deepl_key:
        print("[ERROR] Missing DEEPL_API_KEY env. Please set repository secret DEEPL_API_KEY.")
        sys.exit(2)

    translator = deepl.Translator(deepl_key)

    prs = Presentation(src_path)

    for si, slide in enumerate(prs.slides, start=1):
        # Text frames and tables
        for shape in slide.shapes:
            # Text frames
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                original = shape.text
                translated = translate_text(translator, original)
                if translated != original:
                    shape.text = translated
            # Tables
            if hasattr(shape, "has_table") and shape.has_table:
                tbl = shape.table
                for r in range(len(tbl.rows)):
                    for c in range(len(tbl.columns)):
                        cell = tbl.cell(r, c)
                        original = cell.text
                        translated = translate_text(translator, original)
                        if translated != original:
                            cell.text = translated
            # Pictures (OCR optional)
            if enable_ocr and shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = shape.image.blob
                    ocr_txt = ocr_image_to_text(blob)
                    if ocr_txt and not is_chinese_text(ocr_txt):
                        translated = translate_text(translator, ocr_txt)
                        # Append OCR/translation to alternative text for reference
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
                    translated = translate_text(translator, original)
                    if translated != original:
                        notes_slide.notes_text_frame.text = translated
        except Exception as e:
            print(f"[WARN] Notes processing failed on slide {si}: {e}")

    prs.save(out_path)
    print(f"[INFO] Saved translated PPT: {out_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Translate non-Chinese PPT text to Simplified Chinese")
    parser.add_argument("--source", default="01计算机设计导论.pptx", help="Source PPTX path")
    parser.add_argument("--output", default="01计算机设计导论_zh-CN.pptx", help="Output PPTX path")
    parser.add_argument("--enable-ocr", action="store_true", help="Enable OCR for pictures")
    args = parser.parse_args()

    process_presentation(args.source, args.output, enable_ocr=args.enable_ocr)
