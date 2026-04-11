"""
drivemap.py — Recursively maps Google Drive "Shared with me" folders matching
"NEW CFJ RESOURCE FOLDER Rev 2024" (case-insensitive exact match), reads
"1 LINKS" xlsx files inline, detects duplicates, and outputs
drive_map.md + drive_map.html.
"""

import io
import os
import re
import sys
import html as html_lib
from collections import Counter
from datetime import datetime

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("WARNING: openpyxl not installed — xlsx content will not be read.")

# ── Constants ──────────────────────────────────────────────────────────────────
SCOPES      = ["https://www.googleapis.com/auth/drive.readonly"]
CREDS_FILE  = "credentials.json"
TOKEN_FILE  = "token.json"
MD_OUTPUT   = "drive_map.md"
HTML_OUTPUT = "drive_map.html"
TARGET_NAME = "NEW CFJ RESOURCE FOLDER Rev 2024"
FOLDER_MIME = "application/vnd.google-apps.folder"
XLSX_MIME   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FIELDS      = "nextPageToken, files(id, name, mimeType, webViewLink, modifiedTime)"

MIME_DESCRIPTIONS = {
    FOLDER_MIME:                                                                   "Google Drive folder",
    "application/vnd.google-apps.document":                                        "Google Docs document",
    "application/vnd.google-apps.spreadsheet":                                     "Google Sheets spreadsheet",
    "application/vnd.google-apps.presentation":                                    "Google Slides presentation",
    "application/vnd.google-apps.form":                                            "Google Form",
    "application/vnd.google-apps.drawing":                                         "Google Drawing",
    "application/vnd.google-apps.script":                                          "Google Apps Script",
    "application/vnd.google-apps.site":                                            "Google Site",
    "application/pdf":                                                             "PDF document",
    "image/jpeg":                                                                  "JPEG image",
    "image/png":                                                                   "PNG image",
    "image/gif":                                                                   "GIF image",
    "image/svg+xml":                                                               "SVG image",
    "video/mp4":                                                                   "MP4 video",
    "video/quicktime":                                                             "QuickTime video",
    "audio/mpeg":                                                                  "MP3 audio",
    "text/plain":                                                                  "Plain text file",
    "text/csv":                                                                    "CSV data file",
    "application/zip":                                                             "ZIP archive",
    "application/x-zip-compressed":                                               "ZIP archive",
    "application/msword":                                                          "Word document",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document":    "Word document",
    "application/vnd.ms-excel":                                                    "Excel spreadsheet",
    XLSX_MIME:                                                                     "Excel spreadsheet",
    "application/vnd.ms-powerpoint":                                               "PowerPoint presentation",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation":  "PowerPoint presentation",
}


# ── Lookup helpers ─────────────────────────────────────────────────────────────

def mime_label(mime: str) -> str:
    labels = {
        FOLDER_MIME:                                                                   "Folder",
        "application/vnd.google-apps.document":                                        "Google Doc",
        "application/vnd.google-apps.spreadsheet":                                     "Google Sheet",
        "application/vnd.google-apps.presentation":                                    "Google Slides",
        "application/vnd.google-apps.form":                                            "Google Form",
        "application/vnd.google-apps.drawing":                                         "Google Drawing",
        "application/vnd.google-apps.script":                                          "Apps Script",
        "application/vnd.google-apps.site":                                            "Google Site",
        "application/pdf":                                                             "PDF",
        "image/jpeg":                                                                  "JPEG",
        "image/png":                                                                   "PNG",
        "image/gif":                                                                   "GIF",
        "image/svg+xml":                                                               "SVG",
        "video/mp4":                                                                   "MP4 Video",
        "video/quicktime":                                                             "QuickTime Video",
        "audio/mpeg":                                                                  "MP3 Audio",
        "text/plain":                                                                  "Text",
        "text/csv":                                                                    "CSV",
        "application/zip":                                                             "ZIP",
        "application/msword":                                                          "Word Doc",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document":    "Word Doc",
        "application/vnd.ms-excel":                                                    "Excel",
        XLSX_MIME:                                                                     "Excel",
        "application/vnd.ms-powerpoint":                                               "PowerPoint",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation":  "PowerPoint",
    }
    return labels.get(mime, mime.split("/")[-1].replace("vnd.", "").replace(".", " ").title())


def badge_class(mime: str) -> str:
    sheets = {
        "application/vnd.google-apps.spreadsheet",
        "application/vnd.ms-excel",
        XLSX_MIME,
    }
    docs = {
        "application/vnd.google-apps.document",
        "application/msword",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    slides = {
        "application/vnd.google-apps.presentation",
        "application/vnd.ms-powerpoint",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }
    audio = {
        "audio/mpeg", "audio/mp4", "audio/x-m4a", "audio/wav",
        "audio/x-wav", "audio/aac", "audio/ogg", "audio/flac",
        "audio/webm", "audio/x-ms-wma",
    }
    if mime in sheets:              return "badge-sheet"
    if mime in docs:                return "badge-doc"
    if mime == "application/pdf":   return "badge-pdf"
    if mime in slides:              return "badge-slides"
    if mime == FOLDER_MIME:         return "badge-folder"
    if mime in audio:               return "badge-audio"
    return "badge-other"


AUDIO_MIMES = {
    "audio/mpeg", "audio/mp4", "audio/x-m4a", "audio/wav",
    "audio/x-wav", "audio/aac", "audio/ogg", "audio/flac",
    "audio/webm", "audio/x-ms-wma",
}
VIDEO_MIMES = {"video/mp4", "video/quicktime", "video/webm", "video/x-msvideo"}


def type_key(mime: str) -> str:
    """Short key driving the HTML type-filter buttons."""
    if mime == FOLDER_MIME:         return "folder"
    if mime == "application/pdf":   return "pdf"
    if mime in {
        "application/vnd.google-apps.document",
        "application/msword",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }:                              return "word"
    if mime in {
        "application/vnd.google-apps.spreadsheet",
        "application/vnd.ms-excel",
        XLSX_MIME,
    }:                              return "excel"
    if mime in AUDIO_MIMES:         return "audio"
    if mime in VIDEO_MIMES:         return "video"
    return "other"


# ── Text helpers ───────────────────────────────────────────────────────────────

def clean_text(s: str) -> str:
    """Strip HTML entities, non-breaking spaces, and excess whitespace."""
    s = html_lib.unescape(s)
    s = s.replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"[ \t]+", " ", s).strip()


def fmt_date(iso: str) -> str:
    """Format ISO date string as 'Aug 14, 2020'."""
    if not iso:
        return ""
    try:
        dt = datetime.strptime(iso[:10], "%Y-%m-%d")
        return f"{dt.strftime('%b')} {dt.day}, {dt.year}"
    except ValueError:
        return iso[:10]


def get_description(name: str, mime: str) -> str:
    base = MIME_DESCRIPTIONS.get(mime, mime.split("/")[-1].replace(".", " ").title() + " file")
    stem = os.path.splitext(name)[0].strip()
    if stem and stem.lower() not in ("untitled", "document", "spreadsheet", "presentation"):
        return f"{base} \u2014 {stem}"
    return base


def get_link(file_id: str, mime: str) -> str:
    google_app_bases = {
        "application/vnd.google-apps.document":     "https://docs.google.com/document/d/",
        "application/vnd.google-apps.spreadsheet":  "https://docs.google.com/spreadsheets/d/",
        "application/vnd.google-apps.presentation": "https://docs.google.com/presentation/d/",
        "application/vnd.google-apps.form":         "https://docs.google.com/forms/d/",
        "application/vnd.google-apps.drawing":      "https://docs.google.com/drawings/d/",
    }
    if mime in google_app_bases:
        return f"{google_app_bases[mime]}{file_id}/edit"
    if mime == FOLDER_MIME:
        return f"https://drive.google.com/drive/folders/{file_id}"
    return f"https://drive.google.com/file/d/{file_id}/view"


# ── xlsx reading ───────────────────────────────────────────────────────────────

def read_xlsx_links(service, file_id: str):
    """
    Download and parse an xlsx file.
    Returns:
      None   — download/parse error (silently skipped)
      []     — file has no data rows (only header or empty) → show EMPTY badge
      [...]  — list of row dicts with keys: resource_type, title, link, addl_info
    """
    if not OPENPYXL_AVAILABLE:
        return None
    try:
        buf = io.BytesIO()
        req = service.files().get_media(fileId=file_id)
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)

        wb = openpyxl.load_workbook(buf, read_only=True, data_only=True)
        ws = wb.active
        all_rows = [
            [str(c).strip() if c is not None else "" for c in row]
            for row in ws.iter_rows(values_only=True)
            if any(c is not None for c in row)
        ]
        if not all_rows:
            return []

        raw_headers = [h.lower() for h in all_rows[0]]

        def find_col(*candidates) -> int | None:
            for cand in candidates:
                for i, h in enumerate(raw_headers):
                    if cand in h:
                        return i
            return None

        col_type  = find_col("resource type", "type")
        col_title = find_col("title", "name")
        col_link  = find_col("resource link", "link", "url")
        col_addl  = find_col("add'l", "additional", "addl", "notes", "description")

        n = len(raw_headers)
        if col_type  is None and n > 0: col_type  = 0
        if col_title is None and n > 1: col_title = 1
        if col_link  is None and n > 2: col_link  = 2
        if col_addl  is None and n > 3: col_addl  = 3

        def gcell(row, idx):
            if idx is None or idx >= len(row):
                return ""
            return row[idx]

        results = []
        for row in all_rows[1:]:
            rt   = gcell(row, col_type)
            ttl  = gcell(row, col_title)
            lnk  = gcell(row, col_link)
            addl = gcell(row, col_addl)
            if any([rt, ttl, lnk, addl]):
                results.append({"resource_type": rt, "title": ttl, "link": lnk, "addl_info": addl})

        return results

    except Exception as exc:
        print(f"  WARNING: Could not read xlsx {file_id}: {exc}")
        return None


# ── Drive API ──────────────────────────────────────────────────────────────────

def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                sys.exit(f"ERROR: {CREDS_FILE} not found. Download it from Google Cloud Console.")
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as fh:
            fh.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def list_children(service, folder_id: str) -> list:
    items, page_token = [], None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields=FIELDS,
            pageSize=1000,
            pageToken=page_token,
        ).execute()
        items.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items


def list_all_shared_folders(service) -> list:
    folders, page_token = [], None
    while True:
        resp = service.files().list(
            q="mimeType = 'application/vnd.google-apps.folder' and sharedWithMe = true and trashed = false",
            fields=FIELDS,
            pageSize=1000,
            pageToken=page_token,
        ).execute()
        folders.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return folders


def find_target_folders(service) -> list:
    all_folders = list_all_shared_folders(service)
    print(f"\n[DEBUG] All folders found in 'Shared with me' ({len(all_folders)} total):")
    for f in sorted(all_folders, key=lambda x: x["name"].lower()):
        print(f"  {f['name']}")
    print()
    return [f for f in all_folders if f["name"].lower() == TARGET_NAME.lower()]


# ── Data collection ────────────────────────────────────────────────────────────

def recurse(service, folder_id: str, depth: int,
            rows: list, folder_tree: list,
            parent_anchor: str, parent_folder_id: str,
            folder_path: list):
    """
    DFS walk of folder_id. Children are alphabetised before processing.
    Each row dict carries: row_type, id, depth, name, mime, desc, link,
    modified, anchor, parent_folder_id, folder_path, is_duplicate,
    is_empty_folder, xlsx_rows.
    """
    children = sorted(
        list_children(service, folder_id),
        key=lambda x: x.get("name", "").lower()
    )

    for item in children:
        name     = clean_text(item.get("name", ""))
        mime     = item.get("mimeType", "")
        fid      = item["id"]
        link     = item.get("webViewLink") or get_link(fid, mime)
        desc     = clean_text(get_description(name, mime))
        modified = fmt_date(item.get("modifiedTime", ""))
        anchor   = f"f{fid}" if mime == FOLDER_MIME else ""

        row = {
            "row_type":         "item",
            "id":               fid,
            "depth":            depth,
            "name":             name,
            "mime":             mime,
            "desc":             desc,
            "link":             link,
            "modified":         modified,
            "parent_anchor":    parent_anchor,
            "parent_folder_id": parent_folder_id,
            "anchor":           anchor,
            "folder_path":      list(folder_path),
            "is_duplicate":     False,
            "is_empty_folder":  False,
            "xlsx_rows":        None,
        }
        rows.append(row)

        # Inline-read xlsx files whose names start with "1 LINKS" (case-insensitive)
        if mime == XLSX_MIME and name.lower().startswith("1 links"):
            print(f"    Reading xlsx: {name}")
            row["xlsx_rows"] = read_xlsx_links(service, fid)

        if mime == FOLDER_MIME:
            folder_tree.append({"name": name, "depth": depth, "anchor": anchor})
            folder_path.append(name)
            recurse(service, fid, depth + 1, rows, folder_tree, anchor, fid, folder_path)
            folder_path.pop()


# ── Post-processing ────────────────────────────────────────────────────────────

def detect_duplicates(rows: list):
    """Only mark the 2nd and later occurrences of the same file name (case-insensitive)."""
    counts = Counter(
        r["name"].lower()
        for r in rows
        if r["row_type"] == "item" and r["mime"] != FOLDER_MIME
    )
    seen: set = set()
    for r in rows:
        if r["row_type"] == "item" and r["mime"] != FOLDER_MIME:
            key = r["name"].lower()
            if counts[key] > 1:
                if key in seen:
                    r["is_duplicate"] = True
                else:
                    seen.add(key)


def mark_empty_folders(rows: list):
    """Mark folder rows whose folder_id has no direct non-folder children."""
    folder_ids_with_files = {
        r["parent_folder_id"]
        for r in rows
        if r["row_type"] == "item" and r["mime"] != FOLDER_MIME
    }
    for r in rows:
        if r["row_type"] == "item" and r["mime"] == FOLDER_MIME:
            r["is_empty_folder"] = r["id"] not in folder_ids_with_files


# ── Sidebar tree builder ───────────────────────────────────────────────────────

def _render_tree_node(name: str, anchor: str, children_html: str, is_root: bool = False) -> str:
    esc  = html_lib.escape
    icon = "📂" if is_root else "📁"
    if children_html:
        arrow   = '<span class="tree-arrow">▼</span>'
        onclick = 'onclick="treeToggle(this)"'
        subtree = f'<ul class="tree-children">{children_html}</ul>'
    else:
        arrow   = '<span class="tree-spacer"></span>'
        onclick = ""
        subtree = ""
    return (
        f'<li class="tree-node">'
        f'<div class="tree-item" {onclick}>'
        f'{arrow}'
        f'<a href="#{anchor}" class="tree-link" title="{esc(name)}">{icon} {esc(name)}</a>'
        f'</div>'
        f'{subtree}'
        f'</li>'
    )


def _collect_children(folder_tree: list, idx: int, expected_depth: int):
    """Recursively collect nodes at expected_depth. Returns (html, next_idx)."""
    parts = []
    while idx < len(folder_tree):
        node = folder_tree[idx]
        if node["depth"] < expected_depth:
            break
        if node["depth"] == expected_depth:
            child_html, idx = _collect_children(folder_tree, idx + 1, expected_depth + 1)
            parts.append(_render_tree_node(node["name"], node["anchor"], child_html))
        else:
            idx += 1
    return "".join(parts), idx


def build_sidebar_html(top_folder: dict, folder_tree: list) -> str:
    top_anchor    = f"f{top_folder['id']}"
    children_html, _ = _collect_children(folder_tree, 0, 0)
    top_node      = _render_tree_node(top_folder["name"], top_anchor, children_html, is_root=True)
    return f'<ul class="tree-root">{top_node}</ul>'


# ── Markdown output ────────────────────────────────────────────────────────────

def write_md(rows: list, target_name: str, top_folder: dict):
    esc          = lambda s: s.replace("|", "\\|")
    top_link     = top_folder.get("webViewLink") or get_link(top_folder["id"], FOLDER_MIME)
    top_modified = fmt_date(top_folder.get("modifiedTime", ""))

    lines = [
        "# Google Drive Map\n",
        f"Folder: `{target_name}` (exact match, case-insensitive) from *Shared with me*\n",
        "", "---", "",
        "| File Name | Type | Modified | Description | Link |",
        "|-----------|------|----------|-------------|------|",
        f"| **{target_name}** | Folder | {top_modified} | Top-level matched folder | [Open]({top_link}) |",
    ]

    for r in rows:
        if r["row_type"] != "item":
            continue

        indent      = "&nbsp;&nbsp;" * r["depth"]
        prefix      = "\U0001f4c1 " if r["mime"] == FOLDER_MIME else ""
        dup_marker  = " `[DUP]`" if r["is_duplicate"] else ""
        empty_marker= " `[EMPTY]`" if r.get("is_empty_folder") else ""
        label       = mime_label(r["mime"])

        breadcrumb = ""
        if r["is_duplicate"] and r["folder_path"]:
            breadcrumb = f" *(in: {' / '.join(r['folder_path'])})*"

        lines.append(
            f"| {indent}{prefix}{esc(r['name'])}{dup_marker}{empty_marker}{breadcrumb}"
            f" | {label} | {r['modified']} | {esc(r['desc'])} | [Open]({r['link']}) |"
        )

        if r.get("xlsx_rows") is not None:
            if not r["xlsx_rows"]:
                lines.append(f"| {indent}&nbsp;&nbsp;*EMPTY — no data rows* | | | | |")
            else:
                lines.append(f"| {indent}&nbsp;&nbsp;**Resource Type** \\| **Title** \\| **Link** \\| **Add'l Info** | | | | |")
                for xr in r["xlsx_rows"]:
                    lines.append(
                        f"| {indent}&nbsp;&nbsp;{esc(xr['resource_type'])} \\| "
                        f"{esc(xr['title'])} \\| {esc(xr['link'])} \\| {esc(xr['addl_info'])} | | | | |"
                    )

    with open(MD_OUTPUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


# ── HTML — CSS ─────────────────────────────────────────────────────────────────

CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  font-size: 14px;
  background: #f8fafc;
  color: #1e293b;
  display: flex;
  height: 100vh;
  overflow: hidden;
}

/* ── Sidebar ── */
#sidebar {
  width: 272px;
  min-width: 272px;
  background: #1e293b;
  display: flex;
  flex-direction: column;
  height: 100vh;
  overflow: hidden;
}
#sidebar-header {
  padding: 14px 16px;
  background: #0f172a;
  border-bottom: 1px solid #334155;
  flex-shrink: 0;
}
#sidebar-header h2 {
  font-size: 10.5px;
  font-weight: 700;
  color: #94a3b8;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}
#sidebar-header p {
  font-size: 11.5px;
  color: #e2e8f0;
  margin-top: 4px;
  line-height: 1.4;
}
#sidebar-tree {
  flex: 1;
  overflow-y: auto;
  padding: 6px 0 12px;
}

/* Tree nodes */
.tree-root, .tree-children { list-style: none; padding: 0; margin: 0; }
.tree-children.collapsed { display: none; }
.tree-node { display: block; }
.tree-item {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 4px 10px 4px 8px;
  cursor: default;
  user-select: none;
  transition: background 0.12s;
}
.tree-item:hover { background: #334155; }
.tree-root > .tree-node > .tree-item { padding-left: 8px; }
.tree-children .tree-item { padding-left: 18px; }
.tree-children .tree-children .tree-item { padding-left: 30px; }
.tree-children .tree-children .tree-children .tree-item { padding-left: 42px; }
.tree-children .tree-children .tree-children .tree-children .tree-item { padding-left: 54px; }
.tree-children .tree-children .tree-children .tree-children .tree-children .tree-item { padding-left: 66px; }
.tree-arrow, .tree-spacer {
  font-size: 9px;
  color: #64748b;
  width: 12px;
  flex-shrink: 0;
  cursor: pointer;
  transition: color 0.12s;
}
.tree-item:hover .tree-arrow { color: #94a3b8; }
.tree-link {
  flex: 1;
  text-decoration: none;
  color: #94a3b8;
  font-size: 12px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  transition: color 0.12s;
  min-width: 0;
}
.tree-link:hover { color: #f1f5f9; }
.tree-root-link { color: #e2e8f0; font-weight: 600; font-size: 12.5px; }

/* ── Main area ── */
#main { flex: 1; display: flex; flex-direction: column; overflow: hidden; min-width: 0; }

/* ── Summary bar ── */
#summary {
  background: #fff;
  border-bottom: 1px solid #e2e8f0;
  padding: 11px 20px;
  display: flex;
  gap: 14px;
  align-items: center;
  flex-shrink: 0;
  flex-wrap: wrap;
}
#summary h1 {
  font-size: 14px;
  font-weight: 700;
  color: #0f172a;
  flex: 1;
  min-width: 0;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.stat-card {
  background: #f1f5f9;
  border-radius: 8px;
  padding: 6px 14px;
  text-align: center;
  min-width: 74px;
  transition: background 0.15s, box-shadow 0.15s;
}
.stat-card.clickable { cursor: pointer; }
.stat-card.clickable:hover { background: #e2e8f0; }
.stat-card.active { background: #fef9c3; box-shadow: 0 0 0 2px #f59e0b; }
.stat-num   { font-size: 20px; font-weight: 700; color: #1e293b; line-height: 1.2; }
.stat-label { font-size: 10px; color: #64748b; text-transform: uppercase; letter-spacing: 0.06em; }
.stat-dup  .stat-num { color: #b45309; }
.stat-empty .stat-num { color: #6d28d9; }

/* ── Search bar ── */
#search-bar {
  background: #fff;
  border-bottom: 1px solid #e2e8f0;
  padding: 8px 20px;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  gap: 10px;
}
#search {
  width: 100%;
  max-width: 420px;
  padding: 7px 12px 7px 32px;
  border: 1px solid #cbd5e1;
  border-radius: 6px;
  font-size: 13px;
  outline: none;
  background: #f8fafc url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2.5'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cline x1='21' y1='21' x2='16.65' y2='16.65'/%3E%3C/svg%3E") no-repeat 10px center;
  transition: border-color 0.15s, box-shadow 0.15s;
}
#search:focus {
  border-color: #6366f1;
  box-shadow: 0 0 0 3px rgba(99,102,241,0.12);
  background-color: #fff;
}
#search::placeholder { color: #94a3b8; }
#search-count { font-size: 12px; color: #64748b; white-space: nowrap; }

/* ── Table ── */
#table-wrap { flex: 1; overflow: auto; padding: 12px 20px 24px; }
table {
  width: 100%;
  border-collapse: collapse;
  background: #fff;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
thead th {
  background: #f8fafc;
  color: #64748b;
  font-size: 11px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.07em;
  padding: 10px 12px;
  text-align: left;
  border-bottom: 2px solid #e2e8f0;
  position: sticky;
  top: 0;
  z-index: 2;
}
tbody tr.file-row { border-bottom: 1px solid #f1f5f9; transition: background 0.1s; }
tbody tr.file-row:last-child { border-bottom: none; }
tbody tr.file-row:hover:not(.duplicate) { background: #f8fafc; }
tbody tr.duplicate             { background: #fef9c3; }
tbody tr.duplicate:hover       { background: #fef08a; }
td { padding: 8px 12px; vertical-align: top; }
.name-cell { max-width: 340px; }
.date-cell { white-space: nowrap; color: #64748b; font-size: 12px; width: 104px; }
.desc-cell { color: #475569; font-size: 12.5px; }

/* ── Links ── */
.file-link { color: #2563eb; text-decoration: none; font-weight: 500; }
.file-link:hover { text-decoration: underline; }

/* ── Breadcrumb ── */
.breadcrumb {
  font-size: 11px;
  color: #94a3b8;
  margin-top: 3px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* ── Type badges ── */
.badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 9999px;
  font-size: 11px;
  font-weight: 600;
  white-space: nowrap;
}
.badge-sheet  { background: #dcfce7; color: #166534; }
.badge-doc    { background: #dbeafe; color: #1e40af; }
.badge-pdf    { background: #fee2e2; color: #991b1b; }
.badge-slides { background: #ffedd5; color: #9a3412; }
.badge-folder { background: #e0e7ff; color: #3730a3; }
.badge-audio  { background: #f3e8ff; color: #6b21a8; }
.badge-other  { background: #f1f5f9; color: #475569; }

/* ── Inline badges ── */
.dup-badge {
  display: inline-block;
  margin-left: 6px;
  padding: 1px 5px;
  background: #f59e0b;
  color: #1c1917;
  border-radius: 4px;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.04em;
  vertical-align: middle;
}
.empty-badge {
  display: inline-block;
  margin-left: 6px;
  padding: 1px 5px;
  background: #ddd6fe;
  color: #4c1d95;
  border-radius: 4px;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.04em;
  vertical-align: middle;
}

/* ── xlsx flat resource rows ── */
tr.xlsx-header td {
  background: #e0f2fe;
  color: #0369a1;
  font-size: 10.5px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.06em;
  padding: 5px 12px;
  border-top: 2px solid #bae6fd;
  border-bottom: 1px solid #bae6fd;
}
tr.xlsx-resource-row td {
  background: #f0f9ff;
  color: #0c4a6e;
  font-size: 12.5px;
  padding: 6px 12px;
  border-bottom: 1px solid #e0f2fe;
  vertical-align: top;
}
tr.xlsx-resource-row:last-child td { border-bottom: 2px solid #bae6fd; }
tr.xlsx-empty-row td {
  background: #f8fafc;
  color: #6b7280;
  font-style: italic;
  font-size: 12.5px;
  padding: 7px 12px;
  border-bottom: 2px solid #e2e8f0;
}
.xlsx-link { color: #0369a1; text-decoration: none; }
.xlsx-link:hover { text-decoration: underline; }

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

/* ── Row flash (sidebar nav) ── */
@keyframes rowFlash {
  0%   { background-color: transparent; }
  20%  { background-color: #bfdbfe; }
  100% { background-color: transparent; }
}
tr.row-flash { animation: rowFlash 1.4s ease forwards; }

/* ── Scroll-spy active sidebar item ── */
.tree-item.item-active { background: rgba(99,102,241,0.22); border-left: 2px solid #818cf8; }
.tree-item.item-active .tree-link { color: #e0e7ff !important; font-weight: 600; }

/* ── Sidebar collapse/expand buttons ── */
#sidebar-controls { display: flex; gap: 6px; margin-top: 8px; }
.sidebar-btn {
  flex: 1;
  padding: 3px 0;
  background: #334155;
  color: #94a3b8;
  border: 1px solid #475569;
  border-radius: 4px;
  font-size: 10.5px;
  cursor: pointer;
  transition: background 0.12s, color 0.12s;
}
.sidebar-btn:hover { background: #475569; color: #f1f5f9; }

/* ── Type filter bar ── */
#type-filter {
  background: #fff;
  border-bottom: 1px solid #e2e8f0;
  padding: 7px 20px;
  flex-shrink: 0;
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
  align-items: center;
}
.type-btn {
  padding: 3px 14px;
  border-radius: 9999px;
  border: 1px solid #e2e8f0;
  background: #f8fafc;
  color: #475569;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  transition: background 0.12s, color 0.12s, border-color 0.12s;
  white-space: nowrap;
}
.type-btn:hover { background: #f1f5f9; border-color: #cbd5e1; }
.type-btn[data-type="all"].active    { background: #1e293b;  color: #f1f5f9;  border-color: #1e293b; }
.type-btn[data-type="pdf"].active    { background: #fee2e2;  color: #991b1b;  border-color: #fca5a5; }
.type-btn[data-type="word"].active   { background: #dbeafe;  color: #1e40af;  border-color: #93c5fd; }
.type-btn[data-type="excel"].active  { background: #dcfce7;  color: #166534;  border-color: #86efac; }
.type-btn[data-type="audio"].active  { background: #f3e8ff;  color: #6b21a8;  border-color: #d8b4fe; }
.type-btn[data-type="video"].active  { background: #fef3c7;  color: #92400e;  border-color: #fcd34d; }
.type-btn[data-type="folder"].active { background: #e0e7ff;  color: #3730a3;  border-color: #a5b4fc; }

/* ── Back to top button ── */
#back-top {
  position: fixed;
  bottom: 24px;
  right: 24px;
  width: 38px;
  height: 38px;
  border-radius: 50%;
  background: #6366f1;
  color: #fff;
  border: none;
  font-size: 18px;
  line-height: 38px;
  text-align: center;
  cursor: pointer;
  opacity: 0;
  visibility: hidden;
  transition: opacity 0.2s, visibility 0.2s, background 0.15s, transform 0.15s;
  z-index: 100;
  box-shadow: 0 2px 10px rgba(99,102,241,0.4);
}
#back-top.visible { opacity: 1; visibility: visible; }
#back-top:hover { background: #4f46e5; transform: translateY(-2px); }
"""

# ── HTML — JS ──────────────────────────────────────────────────────────────────

JS = """
(function () {
  var wrap      = document.getElementById('table-wrap');
  var searchEl  = document.getElementById('search');
  var countEl   = document.getElementById('search-count');
  var emptyBtn  = document.getElementById('empty-folder-filter');
  var backTop   = document.getElementById('back-top');
  var allRows   = Array.from(document.querySelectorAll('#file-table tr.filterable'));
  var emptyActive = false;
  var activeType  = 'all';

  // ── Filters ──────────────────────────────────────────────────────────
  function applyFilters() {
    var q = searchEl.value.toLowerCase().trim();

    // First pass: decide which filterable rows pass the non-search filters,
    // then check if any xlsx resource child matches search (so we can boost parent).
    var parentVisible = {};
    allRows.forEach(function (row) {
      var matchEmpty = !emptyActive || row.dataset.emptyFolder === 'true';
      var matchType  = activeType === 'all' || row.dataset.type === activeType;
      var basePass   = matchEmpty && matchType;
      var matchSelf  = !q || (row.dataset.search || '').includes(q);

      // Check if any individual xlsx resource row under this parent matches search
      var childMatch = false;
      if (q && !matchSelf && basePass && row.id) {
        document.querySelectorAll(
          '.xlsx-resource-row[data-xlsx-parent="' + row.id + '"]'
        ).forEach(function (xr) {
          if ((xr.dataset.search || '').includes(q)) childMatch = true;
        });
      }

      parentVisible[row.id] = basePass && (matchSelf || childMatch);
    });

    // Second pass: apply visibility + handle xlsx child rows individually
    var visible = 0;
    allRows.forEach(function (row) {
      var show = parentVisible[row.id] || false;
      row.style.display = show ? '' : 'none';

      if (row.id) {
        // Empty note rows – always follow parent
        document.querySelectorAll(
          '.xlsx-empty-row[data-xlsx-parent="' + row.id + '"]'
        ).forEach(function (xr) { xr.style.display = show ? '' : 'none'; });

        // Resource rows – filter individually; header follows resource visibility
        var anyResource = false;
        document.querySelectorAll(
          '.xlsx-resource-row[data-xlsx-parent="' + row.id + '"]'
        ).forEach(function (xr) {
          var xrShow = show && (!q || (xr.dataset.search || '').includes(q));
          xr.style.display = xrShow ? '' : 'none';
          if (xrShow) anyResource = true;
        });

        // Header row visible only when at least one resource row is visible
        document.querySelectorAll(
          '.xlsx-header[data-xlsx-parent="' + row.id + '"]'
        ).forEach(function (xr) {
          xr.style.display = (show && anyResource) ? '' : 'none';
        });
      }

      if (show) visible++;
    });

    countEl.textContent = (visible === allRows.length)
      ? allRows.length + ' items'
      : visible + ' of ' + allRows.length + ' items';
  }

  countEl.textContent = allRows.length + ' items';
  searchEl.addEventListener('input', applyFilters);

  if (emptyBtn) {
    emptyBtn.addEventListener('click', function () {
      emptyActive = !emptyActive;
      emptyBtn.classList.toggle('active', emptyActive);
      applyFilters();
    });
  }

  // ── Type filter buttons ───────────────────────────────────────────────
  document.querySelectorAll('.type-btn').forEach(function (btn) {
    btn.addEventListener('click', function () {
      activeType = this.dataset.type;
      document.querySelectorAll('.type-btn').forEach(function (b) {
        b.classList.toggle('active', b.dataset.type === activeType);
      });
      applyFilters();
    });
  });

  // ── Sidebar collapse / expand all ────────────────────────────────────
  window.collapseAll = function () {
    document.querySelectorAll('.tree-children').forEach(function (ul) {
      ul.classList.add('collapsed');
    });
    document.querySelectorAll('.tree-arrow').forEach(function (a) {
      a.textContent = '\\u25B6';
    });
  };
  window.expandAll = function () {
    document.querySelectorAll('.tree-children').forEach(function (ul) {
      ul.classList.remove('collapsed');
    });
    document.querySelectorAll('.tree-arrow').forEach(function (a) {
      a.textContent = '\\u25BC';
    });
  };

  // ── Sidebar nav: click → smooth scroll + flash ───────────────────────
  document.querySelectorAll('.tree-link').forEach(function (link) {
    link.addEventListener('click', function (e) {
      var href = this.getAttribute('href');
      if (!href) return;
      var target = document.getElementById(href.slice(1));
      if (!target) return;
      e.preventDefault();
      var rowRect  = target.getBoundingClientRect();
      var wrapRect = wrap.getBoundingClientRect();
      wrap.scrollTo({ top: wrap.scrollTop + rowRect.top - wrapRect.top - 80, behavior: 'smooth' });
      target.classList.remove('row-flash');
      void target.offsetWidth;
      target.classList.add('row-flash');
      setTimeout(function () { target.classList.remove('row-flash'); }, 1400);
    });
  });

  // ── Back to top ──────────────────────────────────────────────────────
  if (backTop) {
    backTop.addEventListener('click', function () {
      wrap.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }

  // ── Scroll spy ───────────────────────────────────────────────────────
  var folderRows = Array.from(document.querySelectorAll('#file-table tr[id^="f"]'));
  var spyLinks = {};
  document.querySelectorAll('.tree-link').forEach(function (link) {
    var href = link.getAttribute('href');
    if (href) spyLinks[href.slice(1)] = link;
  });

  function updateSpy() {
    var scrollMid = wrap.scrollTop + 120;
    var activeId = null;
    folderRows.forEach(function (row) {
      if (row.offsetTop <= scrollMid) activeId = row.id;
    });
    Object.keys(spyLinks).forEach(function (id) {
      var link   = spyLinks[id];
      var active = (id === activeId);
      link.classList.toggle('spy-active', active);
      link.closest('.tree-item').classList.toggle('item-active', active);
    });
  }

  wrap.addEventListener('scroll', function () {
    if (backTop) backTop.classList.toggle('visible', wrap.scrollTop > 200);
    updateSpy();
  });

  // ── Sidebar collapsible tree ─────────────────────────────────────────
  window.treeToggle = function (el) {
    var li       = el.closest('.tree-node');
    var children = li.querySelector(':scope > .tree-children');
    var arrow    = el.querySelector('.tree-arrow');
    if (!children) return;
    var isOpen = !children.classList.contains('collapsed');
    children.classList.toggle('collapsed', isOpen);
    arrow.textContent = isOpen ? '\\u25B6' : '\\u25BC';
  };
})();
"""


# ── HTML output ────────────────────────────────────────────────────────────────

def write_html(rows: list, folder_tree: list, top_folder: dict, counts: dict):
    esc          = html_lib.escape
    top_link     = top_folder.get("webViewLink") or get_link(top_folder["id"], FOLDER_MIME)
    top_modified = fmt_date(top_folder.get("modifiedTime", ""))
    top_anchor   = f"f{top_folder['id']}"
    title_esc    = esc(TARGET_NAME)

    sidebar = build_sidebar_html(top_folder, folder_tree)

    # ── Table rows ──
    table_parts = [
        f'<tr id="row-{top_anchor}" class="file-row filterable"'
        f' data-search="{esc(top_folder["name"].lower())} folder" data-empty-folder="false" data-type="folder">'
        f'<td class="name-cell"><a href="{top_link}" target="_blank" class="file-link" title="{esc(top_link)}">'
        f'📂 {esc(top_folder["name"])}</a></td>'
        f'<td><span class="badge badge-folder">Folder</span></td>'
        f'<td class="date-cell">{top_modified}</td>'
        f'<td class="desc-cell">Top-level matched folder</td>'
        f'</tr>'
    ]

    for r in rows:
        if r["row_type"] != "item":
            continue

        fid        = r["id"]
        # Folder rows use the anchor id ("f{fid}") so sidebar href="#f{fid}" resolves.
        # File rows use "row-{fid}".
        row_id     = r["anchor"] if r["anchor"] else f"row-{fid}"
        indent_px  = r["depth"] * 18 + 12
        dup_class  = " duplicate" if r["is_duplicate"] else ""
        prefix     = "📁 " if r["mime"] == FOLDER_MIME else ""
        bc         = badge_class(r["mime"])
        label      = mime_label(r["mime"])
        empty_fld  = "true" if r.get("is_empty_folder") else "false"
        tkey       = type_key(r["mime"])
        search_val = esc(f"{r['name']} {label} {r['desc']}".lower())

        # inline badges
        dup_badge   = '<span class="dup-badge">DUPLICATE</span>'  if r["is_duplicate"]        else ""
        empty_badge = '<span class="empty-badge">EMPTY</span>'    if r.get("is_empty_folder") else ""

        # breadcrumb (duplicate rows only)
        breadcrumb_html = ""
        if r["is_duplicate"] and r["folder_path"]:
            path_str = " / ".join(esc(p) for p in r["folder_path"])
            breadcrumb_html = f'<div class="breadcrumb">📁 {path_str}</div>'

        table_parts.append(
            f'<tr id="{row_id}" class="file-row filterable{dup_class}"'
            f' data-search="{search_val}" data-empty-folder="{empty_fld}" data-type="{tkey}">'
            f'<td class="name-cell" style="padding-left:{indent_px}px">'
            f'<a href="{r["link"]}" target="_blank" class="file-link" title="{esc(r["link"])}">{prefix}{esc(r["name"])}</a>'
            f'{dup_badge}{empty_badge}'
            f'{breadcrumb_html}'
            f'</td>'
            f'<td><span class="badge {bc}">{esc(label)}</span></td>'
            f'<td class="date-cell">{r["modified"]}</td>'
            f'<td class="desc-cell">{esc(r["desc"])}</td>'
            f'</tr>'
        )

        # xlsx sub-rows (linked to parent via data-xlsx-parent)
        if r.get("xlsx_rows") is not None:
            res_indent = indent_px + 24
            if not r["xlsx_rows"]:
                # Single "empty" annotation row
                table_parts.append(
                    f'<tr class="xlsx-empty-row" data-xlsx-parent="{row_id}">'
                    f'<td style="padding-left:{res_indent}px" colspan="4">'
                    f'No data rows found in this spreadsheet.</td></tr>'
                )
            else:
                # Column header row (not filterable — follows parent)
                table_parts.append(
                    f'<tr class="xlsx-header" data-xlsx-parent="{row_id}">'
                    f'<td style="padding-left:{res_indent}px">Resource Type</td>'
                    f'<td>Title</td>'
                    f'<td colspan="2">Add\'l Info</td>'
                    f'</tr>'
                )
                # One flat <tr> per resource — individually searchable
                for xr in r["xlsx_rows"]:
                    lnk  = xr["link"]
                    _url = ("https://" + lnk) if lnk.startswith("www.") else lnk
                    is_url = _url.startswith("http://") or _url.startswith("https://")
                    ttl    = xr["title"] or _url
                    title_html = (
                        f'<a href="{esc(_url)}" target="_blank" class="xlsx-link">{esc(ttl)}</a>'
                        if is_url else esc(xr["title"])
                    )
                    xr_search = esc(
                        f"{xr['resource_type']} {xr['title']} {xr['addl_info']}".lower()
                    )
                    table_parts.append(
                        f'<tr class="xlsx-resource-row" data-xlsx-parent="{row_id}"'
                        f' data-search="{xr_search}">'
                        f'<td style="padding-left:{res_indent}px">{esc(xr["resource_type"])}</td>'
                        f'<td>{title_html}</td>'
                        f'<td colspan="2">{esc(xr["addl_info"])}</td>'
                        f'</tr>'
                    )

    table_html = "\n    ".join(table_parts)

    html_out = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Drive Map \u2014 {title_esc}</title>
<style>
{CSS}
</style>
</head>
<body>

<div id="sidebar">
  <div id="sidebar-header">
    <h2>Folder Tree</h2>
    <p>{title_esc}</p>
    <div id="sidebar-controls">
      <button class="sidebar-btn" onclick="collapseAll()">Collapse All</button>
      <button class="sidebar-btn" onclick="expandAll()">Expand All</button>
    </div>
  </div>
  <div id="sidebar-tree">
    {sidebar}
  </div>
</div>

<div id="main">
  <div id="summary">
    <h1>\U0001f4c2 {title_esc}</h1>
    <div class="stat-card">
      <div class="stat-num">{counts['files']}</div>
      <div class="stat-label">Files</div>
    </div>
    <div class="stat-card">
      <div class="stat-num">{counts['folders']}</div>
      <div class="stat-label">Folders</div>
    </div>
    <div class="stat-card stat-dup">
      <div class="stat-num">{counts['duplicates']}</div>
      <div class="stat-label">Duplicates</div>
    </div>
    <div class="stat-card stat-empty clickable" id="empty-folder-filter" title="Click to filter empty folders">
      <div class="stat-num">{counts['empty_folders']}</div>
      <div class="stat-label">Empty Folders</div>
    </div>
  </div>

  <div id="search-bar">
    <input id="search" type="search" placeholder="Search by name, type, or description\u2026" autocomplete="off">
    <span id="search-count"></span>
  </div>

  <div id="type-filter">
    <button class="type-btn active" data-type="all">All</button>
    <button class="type-btn" data-type="pdf">PDF</button>
    <button class="type-btn" data-type="word">Word Doc</button>
    <button class="type-btn" data-type="excel">Excel</button>
    <button class="type-btn" data-type="audio">Audio</button>
    <button class="type-btn" data-type="video">Video</button>
    <button class="type-btn" data-type="folder">Folder</button>
  </div>

  <div id="table-wrap">
    <table>
      <thead>
        <tr>
          <th>File Name</th>
          <th>Type</th>
          <th>Modified</th>
          <th>Description</th>
        </tr>
      </thead>
      <tbody id="file-table">
    {table_html}
      </tbody>
    </table>
  </div>
</div>

<script>
{JS}
</script>

<button id="back-top" title="Back to top">&#8679;</button>

</body>
</html>
"""

    with open(HTML_OUTPUT, "w", encoding="utf-8") as fh:
        fh.write(html_out)


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    print("Authenticating with Google Drive...")
    service = authenticate()

    print("Searching 'Shared with me' for matching folders...")
    targets = find_target_folders(service)

    if not targets:
        print(f"No folder with the exact name '{TARGET_NAME}' found. Nothing written.")
        return

    print(f"Found {len(targets)} matching folder(s). Recursing...")

    rows:        list = []
    folder_tree: list = []
    top_folder        = targets[0]

    for folder in targets:
        print(f"  Mapping: {folder['name']}")
        folder_path = [clean_text(folder["name"])]
        recurse(
            service, folder["id"], 0,
            rows, folder_tree,
            f"f{folder['id']}", folder["id"],
            folder_path,
        )

    detect_duplicates(rows)
    mark_empty_folders(rows)

    item_rows     = [r for r in rows if r["row_type"] == "item"]
    total_files   = sum(1 for r in item_rows if r["mime"] != FOLDER_MIME)
    total_folders = sum(1 for r in item_rows if r["mime"] == FOLDER_MIME)
    total_dupes   = sum(1 for r in item_rows if r["is_duplicate"])
    total_empty   = sum(1 for r in item_rows if r.get("is_empty_folder"))
    counts = {
        "files":         total_files,
        "folders":       total_folders,
        "duplicates":    total_dupes,
        "empty_folders": total_empty,
    }

    write_md(rows, TARGET_NAME, top_folder)
    write_html(rows, folder_tree, top_folder, counts)

    print(f"\nDone!  Files: {total_files}  Folders: {total_folders}  "
          f"Duplicates: {total_dupes}  Empty folders: {total_empty}")
    print(f"  {MD_OUTPUT}")
    print(f"  {HTML_OUTPUT}")


if __name__ == "__main__":
    main()
