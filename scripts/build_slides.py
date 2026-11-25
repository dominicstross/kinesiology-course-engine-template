import argparse
import glob
import io
import json
import os
import re
import shutil
import subprocess
import sys
from typing import Any, Callable, Dict, List, Optional

try:
    from pptx import Presentation  # type: ignore
    from pptx.enum.shapes import MSO_SHAPE_TYPE  # type: ignore
except ImportError:
    Presentation = None  # type: ignore
    MSO_SHAPE_TYPE = None  # type: ignore

try:
    from PIL import Image  # type: ignore
except ImportError:
    Image = None  # type: ignore


def find_repo_root(target_abs: str) -> str:
    curr = target_abs if os.path.isdir(target_abs) else os.path.dirname(target_abs)
    while curr and os.path.dirname(curr) != curr:
        if os.path.basename(curr).lower() == "content":
            return os.path.dirname(curr)
        if os.path.isdir(os.path.join(curr, "content")):
            return curr
        curr = os.path.dirname(curr)
    return os.path.dirname(target_abs)


def get_module_id(pptx_abs: str) -> str:
    base_name = os.path.basename(pptx_abs)
    m = re.match(r"^([0-9]+\.[0-9]+)", base_name)
    return m.group(1) if m else os.path.splitext(base_name)[0]


def inspect_deck_status(pptx_abs: str) -> Dict[str, Any]:
    chapter_dir = os.path.dirname(pptx_abs)
    mod_id = get_module_id(pptx_abs)
    json_path = os.path.join(chapter_dir, f"{mod_id}_slides.json")
    img_dir = os.path.join(chapter_dir, "images", mod_id)
    imgs = sorted(glob.glob(os.path.join(img_dir, "slide_*.*")))

    pptx_mtime = os.path.getmtime(pptx_abs)
    json_exists = os.path.exists(json_path)

    if not json_exists:
        status = "NEW"
    else:
        json_mtime = os.path.getmtime(json_path)
        if pptx_mtime > json_mtime:
            status = "NEEDS UPDATE"
        elif imgs and pptx_mtime > min(os.path.getmtime(p) for p in imgs):
            status = "NEEDS UPDATE"
        else:
            status = "UP TO DATE"

    return {
        "pptx_path": pptx_abs,
        "repo_root": find_repo_root(pptx_abs),
        "chapter": os.path.basename(chapter_dir),
        "mod_id": mod_id,
        "filename": os.path.basename(pptx_abs),
        "status": status,
        "png_count": len(imgs),
        "json_exists": json_exists,
    }


def scan_decks(search_root: str) -> List[Dict[str, Any]]:
    clean_root = os.path.normpath(os.path.abspath(search_root))
    results: List[Dict[str, Any]] = []
    for root, dirs, files in os.walk(clean_root):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d.lower() not in ("node_modules", "__pycache__")]
        for fname in sorted(files):
            if fname.lower().endswith((".pptx", ".ppt")) and not fname.startswith("~$"):
                full_p = os.path.join(root, fname)
                try:
                    results.append(inspect_deck_status(full_p))
                except Exception as exc:
                    print(f"[SCAN WARN] Skipping {full_p}: {exc}")
    return results


def save_optimized_image(
    raw_bytes: bytes,
    out_dir_abs: str,
    base_stem: str,
    original_ext: str = "png",
    img_format: str = "webp",
    max_width: int = 1600,
    quality: int = 85,
) -> str:
    """Resizes and compresses image bytes using Pillow; falls back to raw write if Pillow is absent."""
    fmt = img_format.lower().strip()
    if fmt == "jpeg":
        fmt = "jpg"

    if Image is None:
        ext = original_ext.lower().lstrip(".")
        if ext == "jpeg":
            ext = "jpg"
        out_name = f"{base_stem}.{ext}"
        with open(os.path.join(out_dir_abs, out_name), "wb") as f:
            f.write(raw_bytes)
        return out_name

    with Image.open(io.BytesIO(raw_bytes)) as im:
        # Preserve alpha channel for WebP/PNG; flatten to white only if saving as JPG
        if fmt == "jpg" and im.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im.convert("RGBA"), mask=im.convert("RGBA").split()[3])
            im = bg
        elif im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA" if fmt in ("webp", "png") else "RGB")

        # Downscale proportionally if wider than max_width
        if max_width > 0 and im.width > max_width:
            ratio = max_width / float(im.width)
            new_height = max(1, int(round(im.height * ratio)))
            im = im.resize((max_width, new_height), Image.Resampling.LANCZOS)

        out_name = f"{base_stem}.{fmt}"
        out_path = os.path.join(out_dir_abs, out_name)

        if fmt == "webp":
            im.save(out_path, format="WEBP", quality=quality, method=6)
        elif fmt == "png":
            im.save(out_path, format="PNG", optimize=True)
        else:
            im.save(out_path, format="JPEG", quality=quality, optimize=True)

        return out_name


def export_single_slide_win32(pptx_abs_path: str, slide_index: int) -> bytes:
    """Renders a single slide via PowerPoint COM to a temp PNG and returns its bytes."""
    import tempfile
    import win32com.client  # type: ignore

    fd, tmp_png = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        presentation = ppt_app.Presentations.Open(
            pptx_abs_path, ReadOnly=True, Untitled=False, WithWindow=False
        )
        try:
            presentation.Slides(slide_index).Export(tmp_png, "PNG", 1920, 1080)
        finally:
            presentation.Close()
        with open(tmp_png, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_png):
            os.remove(tmp_png)


def extract_slide_metadata_and_media(
    pptx_abs_path: str,
    img_dir_abs: str,
    export_media: bool = True,
    img_format: str = "webp",
    max_width: int = 1600,
    quality: int = 85,
) -> List[Dict[str, Any]]:
    if Presentation is None or MSO_SHAPE_TYPE is None:
        raise RuntimeError("python-pptx is not installed. Run: pip install python-pptx")

    if export_media:
        os.makedirs(img_dir_abs, exist_ok=True)
        for old_file in glob.glob(os.path.join(img_dir_abs, "slide_*.*")):
            os.remove(old_file)

    if not os.path.exists(pptx_abs_path) or os.path.getsize(pptx_abs_path) == 0:
        raise RuntimeError(f'File is empty or cloud-only in OneDrive: {pptx_abs_path}')
    try:
        with open(pptx_abs_path, 'rb') as _chk:
            _hdr = _chk.read(4)
        if _hdr != b'PK\x03\x04':
            raise RuntimeError('Not a valid .pptx OpenXML zip (close PowerPoint if open, or re-save legacy .ppt as .pptx).')
    except PermissionError:
        raise RuntimeError('File is currently open/locked in PowerPoint. Save and close PowerPoint first.')
    prs = Presentation(pptx_abs_path)
    slides_meta: List[Dict[str, Any]] = []

    for idx, slide in enumerate(prs.slides, start=1):
        heading = ""
        bullets: List[str] = []

        if slide.shapes.title and slide.shapes.title.has_text_frame:
            heading = slide.shapes.title.text_frame.text.strip()

        pictures = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        has_charts_or_diagrams = any(
            s.shape_type in (MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.GROUP)
            for s in slide.shapes
        )

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if slide.shapes.title and shape == slide.shapes.title:
                continue
            for p in shape.text_frame.paragraphs:
                line = p.text.strip()
                if not line:
                    continue
                if not heading:
                    heading = line
                elif line != heading and line not in bullets:
                    bullets.append(line)

        if not heading:
            heading = f"Slide {idx}"

        raw_notes = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            raw_notes = slide.notes_slide.notes_text_frame.text.strip()

        youtube_url = ""
        force_full_slide = False
        cleaned_note_lines: List[str] = []
        for line in raw_notes.splitlines():
            yt_match = re.match(r"^\s*\[youtube:\s*(https?://[^\]]+)\]\s*$", line, re.I)
            if yt_match:
                youtube_url = yt_match.group(1).strip()
            elif re.match(r"^\s*\[full-slide\]\s*$", line, re.I):
                force_full_slide = True
            else:
                cleaned_note_lines.append(line)

        transcript = "\n".join(cleaned_note_lines).strip()
        image_filename = ""
        stem = f"slide_{idx:02d}"

        if not youtube_url:
            if len(pictures) == 1 and not has_charts_or_diagrams and not force_full_slide:
                pic = pictures[0]
                if export_media:
                    image_filename = save_optimized_image(
                        pic.image.blob,
                        img_dir_abs,
                        stem,
                        original_ext=pic.image.ext,
                        img_format=img_format,
                        max_width=max_width,
                        quality=quality,
                    )
                else:
                    image_filename = f"{stem}.{img_format}"
            elif len(pictures) > 1 or has_charts_or_diagrams or force_full_slide:
                if export_media:
                    raw_png = export_single_slide_win32(pptx_abs_path, idx)
                    image_filename = save_optimized_image(
                        raw_png,
                        img_dir_abs,
                        stem,
                        original_ext="png",
                        img_format=img_format,
                        max_width=max_width,
                        quality=quality,
                    )
                else:
                    image_filename = f"{stem}.{img_format}"
                bullets = []

        slides_meta.append({
            "index": idx,
            "heading": heading,
            "bullets": bullets,
            "transcript": transcript,
            "youtube": youtube_url,
            "image_filename": image_filename,
        })

    return slides_meta


def should_export_images(
    pptx_abs: str, img_dir_abs: str, force: bool, text_only: bool
) -> bool:
    if text_only:
        return False
    if force:
        return True
    existing = sorted(glob.glob(os.path.join(img_dir_abs, "slide_*.*")))
    if not existing:
        return True
    pptx_mtime = os.path.getmtime(pptx_abs)
    oldest_mtime = min(os.path.getmtime(p) for p in existing)
    return pptx_mtime > oldest_mtime


def convert_deck(
    pptx_path: str,
    repo_root: Optional[str] = None,
    force_images: bool = False,
    text_only: bool = False,
    img_format: str = "webp",
    max_width: int = 1600,
    quality: int = 85,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    def log(msg: str) -> None:
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    pptx_abs = os.path.abspath(pptx_path)
    resolved_root = repo_root or find_repo_root(pptx_abs)
    chapter_dir = os.path.dirname(pptx_abs)
    chapter_rel = os.path.relpath(chapter_dir, resolved_root).replace("\\", "/")

    base_name = os.path.basename(pptx_abs)
    mod_id = get_module_id(pptx_abs)

    img_dir_abs = os.path.join(chapter_dir, "images", mod_id)
    audio_dir_abs = os.path.join(chapter_dir, "audio", mod_id)
    out_json_path = os.path.join(chapter_dir, f"{mod_id}_slides.json")

    do_export = should_export_images(pptx_abs, img_dir_abs, force_images, text_only)
    slides_meta = extract_slide_metadata_and_media(
        pptx_abs,
        img_dir_abs,
        export_media=do_export,
        img_format=img_format,
        max_width=max_width,
        quality=quality,
    )
    total_slides = len(slides_meta)

    log(f"[PROCESSING] {base_name} (Module {mod_id}, {total_slides} slides)")
    if do_export:
        exported_files = glob.glob(os.path.join(img_dir_abs, "slide_*.*"))
        total_kb = sum(os.path.getsize(p) for p in exported_files) / 1024.0
        fmt_note = img_format.upper() if Image is not None else "RAW (install Pillow for WebP)"
        log(f"  [IMAGES] Saved {len(exported_files)} {fmt_note} visual(s) ({total_kb:.1f} KB total) -> {chapter_rel}/images/{mod_id}/")
    else:
        log("  [IMAGES] Up-to-date (skipped media extraction)")

    existing_slides_by_idx: Dict[int, Dict[str, Any]] = {}
    existing_deck_title = ""
    if os.path.exists(out_json_path):
        try:
            with open(out_json_path, "r", encoding="utf-8") as ef:
                edata = json.load(ef)
            existing_deck_title = str(edata.get("title", ""))
            for i, s in enumerate(edata.get("slides", []), start=1):
                existing_slides_by_idx[i] = s
        except Exception:
            pass

    deck_title = existing_deck_title or f"Module {mod_id} Slide Presentation"
    content_json_path = os.path.join(chapter_dir, f"{mod_id}_content.json")
    if not existing_deck_title and os.path.exists(content_json_path):
        try:
            with open(content_json_path, "r", encoding="utf-8") as cf:
                cdata = json.load(cf)
            meta = cdata.get("meta", {}) if isinstance(cdata, dict) else {}
            found_title = meta.get("moduleTitle") or meta.get("title") or cdata.get("title")
            if found_title:
                deck_title = f"{mod_id}: {found_title}"
        except Exception:
            pass

    slides_output: List[Dict[str, Any]] = []
    notes_count = 0
    audio_count = 0

    for meta in slides_meta:
        idx = int(meta["index"])
        prev = existing_slides_by_idx.get(idx, {})

        audio_rel = ""
        for ext in ("wav", "mp3", "m4a", "ogg"):
            candidate = os.path.join(audio_dir_abs, f"slide_{idx:02d}.{ext}")
            if os.path.exists(candidate):
                audio_rel = f"{chapter_rel}/audio/{mod_id}/slide_{idx:02d}.{ext}"
                break

        yt_val = str(meta["youtube"] or prev.get("youtube", ""))
        desc_val = str(prev.get("description", ""))

        slide_entry: Dict[str, Any] = {
            "heading": meta["heading"],
            "description": desc_val,
            "bullets": meta["bullets"],
        }

        if yt_val:
            slide_entry["youtube"] = yt_val
        elif meta["image_filename"]:
            img_file = meta["image_filename"]
            if not os.path.exists(os.path.join(img_dir_abs, img_file)):
                matches = glob.glob(os.path.join(img_dir_abs, f"slide_{idx:02d}.*"))
                if matches:
                    img_file = os.path.basename(matches[0])
            if os.path.exists(os.path.join(img_dir_abs, img_file)):
                slide_entry["image"] = f"{chapter_rel}/images/{mod_id}/{img_file}"
                slide_entry["alt"] = str(prev.get("alt") or meta["heading"])

        if audio_rel:
            slide_entry["audio"] = audio_rel
            audio_count += 1
        if meta["transcript"]:
            slide_entry["transcript"] = meta["transcript"]
            notes_count += 1

        slides_output.append(slide_entry)

    payload = {
        "title": deck_title,
        "slides": slides_output,
    }
    with open(out_json_path, "w", encoding="utf-8") as out_f:
        json.dump(payload, out_f, indent=2, ensure_ascii=False)

    log(f"  [JSON]   Wrote {mod_id}_slides.json ({notes_count}/{total_slides} notes, {audio_count}/{total_slides} audio)")

    return {
        "mod_id": mod_id,
        "total_slides": total_slides,
        "notes_count": notes_count,
        "audio_count": audio_count,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Universal .pptx to _slides.json converter for Kinesiology OLM repos")
    parser.add_argument("target", help="Path to a chapter directory or a single .pptx file")
    parser.add_argument("--force-images", action="store_true", help="Force re-extracting slide visuals")
    parser.add_argument("--text-only", action="store_true", help="Only update JSON text/notes/audio without extracting visuals")
    parser.add_argument("--format", default="webp", choices=["webp", "png", "jpg"], help="Output image format (default: webp)")
    parser.add_argument("--max-width", type=int, default=1600, help="Max image width in pixels (default: 1600)")
    parser.add_argument("--quality", type=int, default=85, help="Compression quality 1-100 for WebP/JPG (default: 85)")
    args = parser.parse_args()

    target_abs = os.path.abspath(args.target)
    if not os.path.exists(target_abs):
        print(f"[ERROR] Target path not found: {target_abs}")
        sys.exit(1)

    repo_root = find_repo_root(target_abs)
    print(f"[REPO ROOT] {repo_root}")

    if os.path.isfile(target_abs) and target_abs.lower().endswith(".pptx"):
        pptx_files = [target_abs]
    else:
        pptx_files = sorted(glob.glob(os.path.join(target_abs, "*.pptx")))

    if not pptx_files:
        print(f"[ERROR] No .pptx files found at {target_abs}")
        sys.exit(1)

    for pptx in pptx_files:
        convert_deck(
            pptx,
            repo_root,
            force_images=args.force_images,
            text_only=args.text_only,
            img_format=args.format,
            max_width=args.max_width,
            quality=args.quality,
        )

    print("\n[ALL COMPLETE] Conversion finished.")


if __name__ == "__main__":
    main()
