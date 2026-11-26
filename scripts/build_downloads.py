import json
import os
import re
import sys
from typing import Any, Dict, List

ALLOWED_EXTS = {".pdf", ".docx", ".doc", ".xlsx", ".pptx", ".zip"}


def human_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.0f} KB"
    return f"{num_bytes / (1024 * 1024):.1f} MB"


def clean_title(filename: str) -> str:
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r"^[0-9]+(?:\.[0-9]+)?[_\-\s]+", "", stem)
    return stem.replace("_", " ").replace("-", " ").strip() or filename


def parse_links_md(md_path: str) -> List[Dict[str, str]]:
    if not os.path.exists(md_path):
        return []
    links: List[Dict[str, str]] = []
    current_scope = "general"
    with open(md_path, "r", encoding="utf-8") as f:
        for raw in f:
            line = raw.strip()
            h_match = re.match(r"^##+\s+(.+)$", line)
            if h_match:
                heading = h_match.group(1).strip()
                mod_m = re.match(r"^([0-9]+\.[0-9]+|ch[0-9]+)", heading, re.I)
                current_scope = mod_m.group(1).lower() if mod_m else heading
                continue
            l_match = re.match(r"^[-*]\s+\[([^\]]+)\]\((https?://[^)]+)\)(?:\s*\|\s*(.+))?$", line)
            if l_match:
                links.append({
                    "title": l_match.group(1).strip(),
                    "url": l_match.group(2).strip(),
                    "description": (l_match.group(3) or "").strip(),
                    "scope": current_scope,
                })
    return links


def build_downloads(repo_root: str) -> str:
    dl_dir = os.path.join(repo_root, "downloads")
    os.makedirs(dl_dir, exist_ok=True)

    files_out: List[Dict[str, Any]] = []
    for root, _, files in os.walk(dl_dir):
        for fname in sorted(files):
            if fname.startswith("~$") or fname == "manifest.json":
                continue
            ext = os.path.splitext(fname)[1].lower()
            if ext not in ALLOWED_EXTS:
                continue
            full_p = os.path.join(root, fname)
            rel_p = os.path.relpath(full_p, repo_root).replace("\\", "/")
            sub = os.path.relpath(root, dl_dir).replace("\\", "/")
            scope = "general" if sub == "." else sub.lower()
            mod_m = re.match(r"^([0-9]+\.[0-9]+)", fname)
            if mod_m:
                scope = mod_m.group(1)

            files_out.append({
                "title": clean_title(fname),
                "filename": fname,
                "path": rel_p,
                "ext": ext.lstrip(".").upper(),
                "size": human_size(os.path.getsize(full_p)),
                "scope": scope,
            })

    links_out = parse_links_md(os.path.join(dl_dir, "links.md"))
    manifest_path = os.path.join(dl_dir, "manifest.json")
    with open(manifest_path, "w", encoding="utf-8") as out:
        json.dump({"files": files_out, "links": links_out}, out, indent=2, ensure_ascii=False)

    print(f"[DOWNLOADS] Wrote {manifest_path} ({len(files_out)} file(s), {len(links_out)} link(s))")
    return manifest_path


if __name__ == "__main__":
    target = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.getcwd()
    build_downloads(target)
