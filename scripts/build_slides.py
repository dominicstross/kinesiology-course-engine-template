import argparse
import glob
import io
import json
import os
import re
import sys
from typing import Any, Callable, Dict, List, Optional, Tuple

try:
    from pptx import Presentation  # type: ignore
    from pptx.enum.shapes import MSO_SHAPE_TYPE  # type: ignore
except ImportError:
    Presentation = None  # type: ignore
    MSO_SHAPE_TYPE = None  # type: ignore

try:
    import fitz  # type: ignore  # PyMuPDF for direct NotebookLM PDF slide rendering
except ImportError:
    fitz = None  # type: ignore

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


def get_module_id(deck_abs: str) -> str:
    base_name = os.path.basename(deck_abs)
    m = re.match(r"^([0-9]+\.[0-9]+)", base_name)
    return m.group(1) if m else os.path.splitext(base_name)[0]


def is_pdf_deck(deck_abs: str) -> bool:
    return deck_abs.lower().endswith(".pdf")


def get_pdf_page_count(pdf_abs: str) -> int:
    if fitz is None or not os.path.exists(pdf_abs) or os.path.getsize(pdf_abs) == 0:
        return 0
    try:
        with fitz.open(pdf_abs) as doc:
            return len(doc)
    except Exception:
        return 0


def count_companion_notes(chapter_dir: str, mod_id: str) -> Tuple[int, int]:
    notes_json_path = os.path.join(chapter_dir, f"{mod_id}_notes.json")
    if not os.path.exists(notes_json_path):
        return (0, 0)
    try:
        with open(notes_json_path, "r", encoding="utf-8") as nf:
            payload = json.load(nf)
        items = payload if isinstance(payload, list) else payload.get("slides", [])
        notes = sum(1 for item in items if str(item.get("transcript", "")).strip())
        return (notes, len(items))
    except Exception:
        return (0, 0)


def count_pptx_notes(pptx_abs: str) -> Tuple[int, int]:
    if Presentation is None or not os.path.exists(pptx_abs) or os.path.getsize(pptx_abs) == 0:
        return (0, 0)
    try:
        prs = Presentation(pptx_abs)
        total = len(prs.slides)
        notes = 0
        for slide in prs.slides:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                txt = slide.notes_slide.notes_text_frame.text.strip()
                cleaned = [
                    ln for ln in txt.splitlines()
                    if not re.match(r"^\s*\[(youtube|full-slide)[^\]]*\]\s*$", ln, re.I)
                ]
                if "\n".join(cleaned).strip():
                    notes += 1
        return (notes, total)
    except Exception:
        return (0, 0)


def get_source_slide_count(deck_abs: str) -> int:
    if is_pdf_deck(deck_abs):
        return get_pdf_page_count(deck_abs)
    if Presentation is not None:
        try:
            return len(Presentation(deck_abs).slides)
        except Exception:
            return 0
    return 0


def inspect_deck_status(deck_abs: str) -> Dict[str, Any]:
    chapter_dir = os.path.dirname(deck_abs)
    mod_id = get_module_id(deck_abs)
    json_path = os.path.join(chapter_dir, f"{mod_id}_slides.json")
    notes_json_path = os.path.join(chapter_dir, f"{mod_id}_notes.json")
    img_dir = os.path.join(chapter_dir, "images", mod_id)
    imgs = sorted(glob.glob(os.path.join(img_dir, "slide_*.*")))

    deck_mtime = os.path.getmtime(deck_abs)
    notes_mtime = os.path.getmtime(notes_json_path) if os.path.exists(notes_json_path) else 0.0
    latest_src_mtime = max(deck_mtime, notes_mtime)
    json_exists = os.path.exists(json_path)

    notes_count = 0
    total_slides = 0

    def _fresh_counts() -> Tuple[int, int]:
        if is_pdf_deck(deck_abs):
            t_slides = get_pdf_page_count(deck_abs)
            n_count, _ = count_companion_notes(chapter_dir, mod_id)
            return (n_count, t_slides)
        n_count, t_slides = count_pptx_notes(deck_abs)
        if n_count == 0:
            c_notes, _ = count_companion_notes(chapter_dir, mod_id)
            n_count = c_notes
        return (n_count, t_slides)

    if not json_exists:
        status = "NEW"
        notes_count, total_slides = _fresh_counts()
    else:
        json_mtime = os.path.getmtime(json_path)
        if latest_src_mtime > json_mtime or (imgs and deck_mtime > min(os.path.getmtime(p) for p in imgs)):
            status = "NEEDS UPDATE"
            notes_count, total_slides = _fresh_counts()
        else:
            status = "UP TO DATE"
            try:
                with open(json_path, "r", encoding="utf-8") as jf:
                    jdata = json.load(jf)
                slist = jdata.get("slides", [])
                total_slides = len(slist)
                notes_count = sum(1 for s in slist if str(s.get("transcript", "")).strip())
            except Exception:
                notes_count, total_slides = _fresh_counts()

    return {
        "pptx_path": deck_abs,
        "repo_root": find_repo_root(deck_abs),
        "chapter": os.path.basename(chapter_dir),
        "mod_id": mod_id,
        "filename": os.path.basename(deck_abs),
        "status": status,
        "png_count": len(imgs),
        "notes_count": notes_count,
        "total_slides": total_slides,
        "json_exists": json_exists,
    }


def scan_decks(search_root: str) -> List[Dict[str, Any]]:
    clean_root = os.path.normpath(os.path.abspath(search_root))
    results: List[Dict[str, Any]] = []
    for root, dirs, files in os.walk(clean_root):
        dirs[:] = [d for d in dirs if not d.startswith(".") and d.lower() not in ("node_modules", "__pycache__")]
        deck_candidates: Dict[str, str] = {}
        for fname in sorted(files):
            lower = fname.lower()
            if fname.startswith("~$"):
                continue
            if lower.endswith((".pptx", ".ppt", ".pdf")):
                full_p = os.path.join(root, fname)
                mod_id = get_module_id(full_p)
                # Prefer .pptx over .pdf if both exist for the same module ID
                if mod_id not in deck_candidates or lower.endswith((".pptx", ".ppt")):
                    deck_candidates[mod_id] = full_p
        for mod_id in sorted(deck_candidates.keys()):
            full_p = deck_candidates[mod_id]
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
        if fmt == "jpg" and im.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            bg.paste(im.convert("RGBA"), mask=im.convert("RGBA").split()[3])
            im = bg
        elif im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA" if fmt in ("webp", "png") else "RGB")

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


def extract_pdf_metadata_and_media(
    pdf_abs_path: str,
    img_dir_abs: str,
    export_media: bool = True,
    img_format: str = "webp",
    max_width: int = 1600,
    quality: int = 85,
) -> List[Dict[str, Any]]:
    if fitz is None:
        raise RuntimeError("PyMuPDF is not installed. Run: pip install pymupdf")
    if not os.path.exists(pdf_abs_path) or os.path.getsize(pdf_abs_path) == 0:
        raise RuntimeError(f"PDF file is empty or cloud-only in OneDrive: {pdf_abs_path}")

    if export_media:
        os.makedirs(img_dir_abs, exist_ok=True)
        for old_file in glob.glob(os.path.join(img_dir_abs, "slide_*.*")):
            os.remove(old_file)

    slides_meta: List[Dict[str, Any]] = []
    with fitz.open(pdf_abs_path) as doc:
        for idx, page in enumerate(doc, start=1):
            stem = f"slide_{idx:02d}"
            image_filename = f"{stem}.{img_format}"
            if export_media:
                # Render at 2.0x scale (~1920px wide on standard 16:9 widescreen PDFs)
                pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                raw_png = pix.tobytes("png")
                image_filename = save_optimized_image(
                    raw_png,
                    img_dir_abs,
                    stem,
                    original_ext="png",
                    img_format=img_format,
                    max_width=max_width,
                    quality=quality,
                )

            slides_meta.append({
                "index": idx,
                "heading": f"Slide {idx}",
                "bullets": [],
                "transcript": "",
                "youtube": "",
                "image_filename": image_filename,
            })

    return slides_meta


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

    if not os.path.exists(pptx_abs_path) or os.path.getsize(pptx_abs_path) == 0:
        raise RuntimeError(f"File is empty or cloud-only in OneDrive: {pptx_abs_path}")
    try:
        with open(pptx_abs_path, "rb") as _chk:
            _hdr = _chk.read(4)
        if _hdr != b"PK\x03\x04":
            raise RuntimeError("Not a valid .pptx OpenXML zip (close PowerPoint if open, or re-save legacy .ppt as .pptx).")
    except PermissionError:
        raise RuntimeError("File is currently open/locked in PowerPoint. Save and close PowerPoint first.")

    if export_media:
        os.makedirs(img_dir_abs, exist_ok=True)
        for old_file in glob.glob(os.path.join(img_dir_abs, "slide_*.*")):
            os.remove(old_file)

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


def sync_companion_notes_json(deck_abs: str, chapter_dir: str, mod_id: str) -> int:
    notes_json_path = os.path.join(chapter_dir, f"{mod_id}_notes.json")
    if not os.path.exists(notes_json_path):
        return 0
    with open(notes_json_path, "r", encoding="utf-8") as nf:
        payload = json.load(nf)
    items = payload if isinstance(payload, list) else payload.get("slides", [])
    by_idx = {int(item.get("index", i + 1)): item for i, item in enumerate(items)}

    if not is_pdf_deck(deck_abs) and Presentation is not None:
        prs = Presentation(deck_abs)
        updated = 0
        for idx, slide in enumerate(prs.slides, start=1):
            entry = by_idx.get(idx)
            if not entry:
                continue
            transcript = str(entry.get("transcript", "")).strip()
            tf = slide.notes_slide.notes_text_frame
            existing_tags = [
                ln for ln in tf.text.splitlines()
                if re.match(r"^\s*\[(youtube|full-slide)[^\]]*\]\s*$", ln, re.I)
            ]
            if "[full-slide]" not in [t.lower() for t in existing_tags]:
                existing_tags.insert(0, "[full-slide]")
            combined = "\n".join(existing_tags + ([transcript] if transcript else [])).strip()
            if tf.text.strip() != combined:
                tf.text = combined
                updated += 1
        if updated > 0:
            prs.save(deck_abs)
    return len(by_idx)


def should_export_images(
    deck_abs: str,
    img_dir_abs: str,
    out_json_path: str,
    force: bool,
    text_only: bool,
) -> bool:
    if text_only:
        return False
    if force:
        return True
    existing = sorted(glob.glob(os.path.join(img_dir_abs, "slide_*.*")))
    if not existing:
        return True
    if os.path.exists(out_json_path):
        try:
            with open(out_json_path, "r", encoding="utf-8") as jf:
                prev_data = json.load(jf)
            prev_count = len(prev_data.get("slides", []))
            curr_count = get_source_slide_count(deck_abs)
            if prev_count > 0 and curr_count > 0 and curr_count != prev_count:
                return True
            max_img_idx = max(
                int(re.search(r"slide_(\d+)", os.path.basename(p)).group(1))
                for p in existing
                if re.search(r"slide_(\d+)", os.path.basename(p))
            )
            if curr_count > 0 and max_img_idx > curr_count:
                return True
        except Exception:
            pass
    deck_mtime = os.path.getmtime(deck_abs)
    oldest_mtime = min(os.path.getmtime(p) for p in existing)
    return deck_mtime > oldest_mtime


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

    deck_abs = os.path.abspath(pptx_path)
    resolved_root = repo_root or find_repo_root(deck_abs)
    chapter_dir = os.path.dirname(deck_abs)
    chapter_rel = os.path.relpath(chapter_dir, resolved_root).replace("\\", "/")

    base_name = os.path.basename(deck_abs)
    mod_id = get_module_id(deck_abs)

    img_dir_abs = os.path.join(chapter_dir, "images", mod_id)
    audio_dir_abs = os.path.join(chapter_dir, "audio", mod_id)
    out_json_path = os.path.join(chapter_dir, f"{mod_id}_slides.json")
    notes_json_path = os.path.join(chapter_dir, f"{mod_id}_notes.json")

    synced_nblm = sync_companion_notes_json(deck_abs, chapter_dir, mod_id)
    if synced_nblm:
        log(f"  [NBLM]   Loaded {synced_nblm} slide note(s) from {mod_id}_notes.json")

    do_export = should_export_images(deck_abs, img_dir_abs, out_json_path, force_images, text_only)
    if is_pdf_deck(deck_abs):
        slides_meta = extract_pdf_metadata_and_media(
            deck_abs,
            img_dir_abs,
            export_media=do_export,
            img_format=img_format,
            max_width=max_width,
            quality=quality,
        )
    else:
        slides_meta = extract_slide_metadata_and_media(
            deck_abs,
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

    nblm_by_idx: Dict[int, Dict[str, Any]] = {}
    if os.path.exists(notes_json_path):
        try:
            with open(notes_json_path, "r", encoding="utf-8") as nf:
                npayload = json.load(nf)
            nitems = npayload if isinstance(npayload, list) else npayload.get("slides", [])
            for i, item in enumerate(nitems, start=1):
                nblm_by_idx[int(item.get("index", i))] = item
        except Exception:
            pass

    slides_output: List[Dict[str, Any]] = []
    notes_count = 0
    audio_count = 0

    for meta in slides_meta:
        idx = int(meta["index"])
        prev = existing_slides_by_idx.get(idx, {})
        nblm_item = nblm_by_idx.get(idx, {})

        audio_rel = ""
        for ext in ("wav", "mp3", "m4a", "ogg"):
            candidate = os.path.join(audio_dir_abs, f"slide_{idx:02d}.{ext}")
            if os.path.exists(candidate):
                audio_rel = f"{chapter_rel}/audio/{mod_id}/slide_{idx:02d}.{ext}"
                break

        yt_val = str(meta["youtube"] or nblm_item.get("youtube") or prev.get("youtube", ""))
        desc_val = str(prev.get("description", ""))
        final_heading = str(nblm_item.get("heading") or meta["heading"]).strip()
        final_bullets = (
            nblm_item.get("bullets")
            if isinstance(nblm_item.get("bullets"), list) and nblm_item.get("bullets")
            else meta["bullets"]
        )
        final_transcript = str(nblm_item.get("transcript") or meta["transcript"]).strip()

        slide_entry: Dict[str, Any] = {
            "heading": final_heading,
            "description": desc_val,
            "bullets": final_bullets,
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
                slide_entry["alt"] = str(prev.get("alt") or final_heading)

        if audio_rel:
            slide_entry["audio"] = audio_rel
            audio_count += 1
        if final_transcript:
            slide_entry["transcript"] = final_transcript
            notes_count += 1

        slides_output.append(slide_entry)

    payload = {
        "title": deck_title,
        "slides": slides_output,
    }
    with open(out_json_path, "w", encoding="utf-8") as out_f:
        json.dump(payload, out_f, indent=2, ensure_ascii=False)

    log(f"  [JSON]   Wrote {mod_id}_slides.json ({notes_count}/{total_slides} notes, {audio_count}/{total_slides} audio)")

    try:
        from build_downloads import build_downloads
        build_downloads(resolved_root)
    except Exception:
        pass

    return {
        "mod_id": mod_id,
        "total_slides": total_slides,
        "notes_count": notes_count,
        "audio_count": audio_count,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Universal .pptx / .pdf to _slides.json converter for Kinesiology OLM repos")
    parser.add_argument("target", help="Path to a chapter directory or a single .pptx/.pdf file")
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

    if os.path.isfile(target_abs) and target_abs.lower().endswith((".pptx", ".pdf")):
        deck_files = [target_abs]
    else:
        deck_files = sorted(
            glob.glob(os.path.join(target_abs, "*.pptx"))
            + glob.glob(os.path.join(target_abs, "*.pdf"))
        )

    if not deck_files:
        print(f"[ERROR] No .pptx or .pdf files found at {target_abs}")
        sys.exit(1)

    for deck in deck_files:
        convert_deck(
            deck,
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

