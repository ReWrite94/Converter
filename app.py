"""
REST service:  DOCX/PPTX → PDF & extract images + text (deduplicated)
==================================================================

JSON response schema
--------------------
{
  "files": [
      {"name": "…", "size": 12345, "mime": "image/png", "sha1": "<40-hex>", "data": "<base64-utf8>"}
  ],
  "extractedTexts": [
      {"name": "<document stem>", "text": "…plain text with {{<<<sha1>>>}} placeholders…"}
  ]
}

• Every distinct image is returned only once (duplicate detection via SHA‑1).
• In the text, each image position is replaced by its hash placeholder;
  every placeholder hash refers to an entry in "files".
• Images too small to be returned (icons, masks) become {{<<<icon>>>}} (PDF only).

Limits (env vars)
-----------------
MAX_UPLOAD_MB          max. upload size                       (default 50)
MAX_UNZIPPED_MB        max. uncompressed size of DOCX/PPTX    (default 500)
SOFFICE_TIMEOUT        LibreOffice timeout in seconds         (default 120)
SOFFICE_MAX_PARALLEL   concurrent LibreOffice processes       (default 2)
"""
from __future__ import annotations

import base64
import contextlib
import hashlib
import html
import mimetypes
import os
import pathlib
import posixpath
import re
import signal
import subprocess
import threading
import zipfile
from tempfile import TemporaryDirectory
from typing import Dict, List, Set, Tuple

import fitz  # PyMuPDF

PDF_TEXT_BLOCK  = getattr(fitz, "PDF_TEXT_BLOCK", 0)
PDF_IMAGE_BLOCK = getattr(fitz, "PDF_IMAGE_BLOCK", 1)

from fastapi import Depends, FastAPI, File, HTTPException, UploadFile, status
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import JSONResponse

from security import api_key_auth

app = FastAPI(
    title="Doc ⇄ PDF Service (JSON)",
    version="2.5.0",  # PDF placeholders now {{<<<sha1>>>}} like DOCX/PPTX
    # Auth applies to every route; exceptions are listed in security.PUBLIC_PATHS.
    dependencies=[Depends(api_key_auth)],
)

MAX_UPLOAD_BYTES     = int(os.environ.get("MAX_UPLOAD_MB", "50")) * 1024 * 1024
MAX_UNZIPPED_BYTES   = int(os.environ.get("MAX_UNZIPPED_MB", "500")) * 1024 * 1024
SOFFICE_TIMEOUT      = int(os.environ.get("SOFFICE_TIMEOUT", "120"))
SOFFICE_MAX_PARALLEL = int(os.environ.get("SOFFICE_MAX_PARALLEL", "2"))

# Image filter for PDFs: smaller images count as icons
MIN_IMG_W, MIN_IMG_H = 48, 48
MIN_IMG_AREA         = 4_000
MIN_IMG_BYTES        = 1_024

# --------------------------------------------------------------------------- #
# Helper functions                                                             #
# --------------------------------------------------------------------------- #

def to_base64(data: bytes) -> str:
    """bytes → base64‑string (utf‑8)"""
    return base64.b64encode(data).decode()


def sha1(data: bytes) -> str:
    """Return 40‑char SHA‑1 hex digest of *data*."""
    return hashlib.sha1(data).hexdigest()


def placeholder(hash_hex: str) -> str:
    """Text marker for an image: {{<<<sha1>>>}}"""
    return "{{<<<" + hash_hex + ">>>}}"


def file_json(name: str, mime: str, data: bytes, *, hash_hex: str | None = None) -> dict:
    """Compose JSON snippet for one file, optionally adding the hash."""
    d = {
        "name": name,
        "size": len(data),
        "mime": mime,
        "data": to_base64(data),
    }
    if hash_hex:
        d["sha1"] = hash_hex
    return d


async def read_upload(file: UploadFile) -> Tuple[str, bytes]:
    """Return (sanitised filename, content); enforces MAX_UPLOAD_BYTES.

    The client filename is only used for suffix checks and response names –
    never as a path on disk.
    """
    # PureWindowsPath strips directory parts separated by "/" *and* "\"
    name = pathlib.PureWindowsPath(file.filename or "").name
    if not name or name in {".", ".."}:
        raise HTTPException(status.HTTP_400_BAD_REQUEST, "missing filename")
    data = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(
            status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            f"file larger than {MAX_UPLOAD_BYTES // (1024 * 1024)} MB",
        )
    return name, data


# --------------------------------------------------------------------------- #
# LibreOffice conversion                                                       #
# --------------------------------------------------------------------------- #

_soffice_slots = threading.BoundedSemaphore(SOFFICE_MAX_PARALLEL)


def soffice_convert(src: pathlib.Path, outdir: pathlib.Path, *, target_fmt: str) -> pathlib.Path:
    """Convert *src* to *target_fmt* (e.g. pdf/docx) using LibreOffice.

    Blocking – call via run_in_threadpool. Each call uses its own LibreOffice
    profile so parallel conversions don't collide on the profile lock.
    """
    profile = (outdir / "lo-profile").as_uri()
    cmd = [
        "soffice",
        f"-env:UserInstallation={profile}",
        "--headless",
        "--convert-to",
        target_fmt,
        "--outdir",
        str(outdir),
        str(src),
    ]
    with _soffice_slots:
        # own process group so a timeout also kills soffice.bin, not just the wrapper
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, start_new_session=True
        )
        try:
            out, err = proc.communicate(timeout=SOFFICE_TIMEOUT)
        except subprocess.TimeoutExpired:
            with contextlib.suppress(ProcessLookupError):
                os.killpg(proc.pid, signal.SIGKILL)
            proc.communicate()
            raise HTTPException(
                status.HTTP_504_GATEWAY_TIMEOUT,
                f"LibreOffice conversion timed out after {SOFFICE_TIMEOUT}s",
            )

    # LibreOffice exits with 0 on many failures – the output file is the real signal
    dst = outdir / src.with_suffix(f".{target_fmt}").name
    if proc.returncode != 0 or not dst.is_file():
        msg = (err or out or "").strip()[-500:]
        raise HTTPException(
            status.HTTP_422_UNPROCESSABLE_ENTITY,
            f"LibreOffice could not convert the file (exit {proc.returncode}): {msg}",
        )
    return dst


# --------------------------------------------------------------------------- #
# Validation helpers                                                           #
# --------------------------------------------------------------------------- #

def validate_suffix(filename: str, *allowed: str) -> None:
    if not filename.lower().endswith(allowed):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"expected file extension {'/'.join(allowed)}",
        )


def open_office_zip(path: pathlib.Path) -> zipfile.ZipFile:
    """Open DOCX/PPTX as zip; reject corrupt files and zip bombs."""
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise HTTPException(status.HTTP_422_UNPROCESSABLE_ENTITY, "file is not a valid DOCX/PPTX")
    # ZipExtFile never yields more than the declared file_size, so the header sum is a hard cap
    if sum(i.file_size for i in zf.infolist()) > MAX_UNZIPPED_BYTES:
        zf.close()
        raise HTTPException(
            status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            f"uncompressed content larger than {MAX_UNZIPPED_BYTES // (1024 * 1024)} MB",
        )
    return zf

# --------------------------------------------------------------------------- #
# Text extraction helpers                                                      #
# --------------------------------------------------------------------------- #

_SENTINEL = "|||IMG"  # protects hashes while stripping XML tags

_RE_REL    = re.compile(r"Id=\"([^\"]+)\"[^>]*Target=\"([^\"]+)\"")
_RE_BLIP   = re.compile(r"<a:blip[^>]+?r:embed=\"([^\"]+)\"[^>]*/>")
_RE_TAGS   = re.compile(r"<[^>]+>")
_RE_SLD_ID = re.compile(r"<p:sldId\b[^>]*?r:id=\"([^\"]+)\"")
_RE_SLIDE  = re.compile(r"ppt/slides/slide(\d+)\.xml")


def _strip_xml_keep_sentinels(xml: str) -> str:
    no_tags = _RE_TAGS.sub(" ", xml)
    return re.sub(r"\s+", " ", no_tags).strip()


def _sentinel_to_placeholder(text: str) -> str:
    pattern = re.compile(r"\|\|\|IMG([0-9a-f]{40})\|\|\|")
    return pattern.sub(lambda m: placeholder(m.group(1)), text)


def _is_large_image(w: int, h: int, nbytes: int) -> bool:
    return w >= MIN_IMG_W and h >= MIN_IMG_H and w * h >= MIN_IMG_AREA and nbytes >= MIN_IMG_BYTES


def _text_from_rawdict_block(blk: dict) -> str:
    """
    Return the plain text of a RAWDICT or DICT text block.
    Compatible with all PyMuPDF versions (0.x – 1.x).
    """
    # 1) Newer DICT output (PyMuPDF ≥ 1.19) has 'text'
    if blk.get("text"):
        return blk["text"]

    # 2) RAWDICT / DICT: hierarchy  blocks → lines → spans
    if "lines" in blk:
        pieces: list[str] = []
        for line in blk["lines"]:
            for span in line.get("spans", []):
                if span.get("text"):                     # DICT
                    pieces.append(span["text"])
                elif span.get("chars"):                  # RAWDICT ≥ 1.21
                    pieces.append("".join(ch.get("c", "") for ch in span["chars"]))
            pieces.append("\n")                          # keep line breaks
        return "".join(pieces).rstrip("\n")

    # 3) Very old RAWDICT variants (< 1.18) use 'items'
    if "items" in blk:
        return "".join(
            it.get("text", it.get("c", ""))
            for it in blk["items"]
        )

    # Fallback
    return ""

# --------------------------- PDF ------------------------------------------- #

def collect_pdf_images(doc: fitz.Document) -> Tuple[Dict[str, dict], Dict[int, str]]:
    """Collect large images by xref → (unique images by hash, hash by xref)."""
    unique_by_hash: Dict[str, dict] = {}
    hash_by_xref  : Dict[int, str]  = {}
    seen_xrefs    : Set[int]        = set()

    for page in doc:
        for xref, *_ in page.get_images(full=True):
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            info = doc.extract_image(xref)
            if not info:
                continue
            if info["cs-name"] == "DeviceGray" and info["bpc"] == 1:
                continue  # 1‑bit mask images
            if not _is_large_image(info["width"], info["height"], len(info["image"])):
                continue

            hsh = sha1(info["image"])
            hash_by_xref[xref] = hsh
            unique_by_hash.setdefault(hsh, info)

    return unique_by_hash, hash_by_xref


def extract_pdf_text_with_placeholders(
    doc: fitz.Document,
    hash_by_xref: Dict[int, str],
    unique_by_hash: Dict[str, dict],
) -> str:
    """
    Extract PDF text with image placeholders.

        • Images in *hash_by_xref*                 → {{<<<sha1>>>}}
        • Large inline images (no xref) are added
          to *unique_by_hash*                       → {{<<<sha1>>>}}
        • Everything else (icons, masks)           → {{<<<icon>>>}}

    So every hash in the text also appears in the file list.
    """
    parts: List[str] = []

    for page in doc:
        # rawdict image blocks carry no xref → map them via the block number
        xref_by_block = {info["number"]: info["xref"] for info in page.get_image_info(xrefs=True)}

        for blk in page.get_text("rawdict")["blocks"]:
            if blk["type"] == PDF_TEXT_BLOCK:
                parts.append(_text_from_rawdict_block(blk))

            elif blk["type"] == PDF_IMAGE_BLOCK:
                xref = xref_by_block.get(blk["number"], 0)
                hsh: str | None = None
                if xref:
                    hsh = hash_by_xref.get(xref)
                elif _is_large_image(blk["width"], blk["height"], len(blk["image"])):
                    # inline image: only reachable through the text layer
                    hsh = sha1(blk["image"])
                    unique_by_hash.setdefault(hsh, {"ext": blk["ext"], "image": blk["image"]})
                parts.append(placeholder(hsh) if hsh else "{{<<<icon>>>}}")

    return "\n".join(parts).strip()

# --------------------------- DOCX ------------------------------------------ #

def extract_docx_text_with_placeholders(zf: zipfile.ZipFile, img_hash_by_member: Dict[str, str]) -> str:
    try:
        rels_raw = zf.read("word/_rels/document.xml.rels").decode("utf‑8", "ignore")
    except KeyError:
        return ""
    id2target = {rid: tgt for rid, tgt in _RE_REL.findall(rels_raw)}

    lookup: Dict[str, str] = {}
    for member, h in img_hash_by_member.items():
        lookup[member] = h
        lookup[pathlib.Path(member).name] = h
        if "/media/" in member:
            lookup[member.split("/media/", 1)[1]] = h

    doc_xml = zf.read("word/document.xml").decode("utf‑8", "ignore")

    def repl(m: re.Match) -> str:
        rid = m.group(1)
        tgt = id2target.get(rid, "")
        h = lookup.get(tgt) or lookup.get(pathlib.Path(tgt).name)
        return f"{_SENTINEL}{h}{_SENTINEL}" if h else " "

    with_sent = _RE_BLIP.sub(repl, doc_xml)
    text = _strip_xml_keep_sentinels(html.unescape(with_sent))
    return _sentinel_to_placeholder(text)


# --------------------------- PPTX ------------------------------------------ #

def _pptx_slide_paths(zf: zipfile.ZipFile) -> List[str]:
    """Slide XML paths in presentation order (falls back to numeric order)."""
    names = set(zf.namelist())
    try:
        pres_xml  = zf.read("ppt/presentation.xml").decode("utf‑8", "ignore")
        pres_rels = zf.read("ppt/_rels/presentation.xml.rels").decode("utf‑8", "ignore")
    except KeyError:
        pass
    else:
        rid2tgt = dict(_RE_REL.findall(pres_rels))
        ordered = []
        for rid in _RE_SLD_ID.findall(pres_xml):
            tgt = rid2tgt.get(rid)
            if tgt:
                path = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join("ppt", tgt))
                if path in names:
                    ordered.append(path)
        if ordered:
            return ordered

    numbered = [(int(m.group(1)), n) for n in names if (m := _RE_SLIDE.fullmatch(n))]
    return [n for _, n in sorted(numbered)]


def extract_pptx_text_with_placeholders(zf: zipfile.ZipFile, img_hash_by_member: Dict[str, str]) -> str:
    chunks: List[str] = []

    for slide_xml_path in _pptx_slide_paths(zf):
        rels_path = slide_xml_path.replace("/slides/", "/slides/_rels/") + ".rels"
        rid_map: Dict[str, str] = {}
        if rels_path in zf.namelist():
            rels_raw = zf.read(rels_path).decode("utf‑8", "ignore")
            rid_map = {rid: tgt for rid, tgt in _RE_REL.findall(rels_raw)}

        slide_xml = zf.read(slide_xml_path).decode("utf‑8", "ignore")

        def repl(m: re.Match) -> str:
            rid = m.group(1)
            tgt = rid_map.get(rid, "")
            h = (
                img_hash_by_member.get(f"ppt/{tgt}") or
                img_hash_by_member.get(tgt) or
                img_hash_by_member.get(pathlib.Path(tgt).name)
            )
            return f"{_SENTINEL}{h}{_SENTINEL}" if h else " "

        with_sent = _RE_BLIP.sub(repl, slide_xml)
        txt = _strip_xml_keep_sentinels(html.unescape(with_sent))
        chunks.append(_sentinel_to_placeholder(txt))
        chunks.append("\n")
    return "\n".join(chunks).strip()

# --------------------------------------------------------------------------- #
# Blocking workers (run in threadpool, never on the event loop)                #
# --------------------------------------------------------------------------- #

def _convert_to_pdf(name: str, data: bytes, suffix: str) -> dict:
    with TemporaryDirectory() as td:
        src = pathlib.Path(td) / f"input{suffix}"
        src.write_bytes(data)
        pdf_bytes = soffice_convert(src, src.parent, target_fmt="pdf").read_bytes()
    pdf_name = f"{pathlib.Path(name).stem}.pdf"
    return {
        "files": [file_json(pdf_name, "application/pdf", pdf_bytes, hash_hex=sha1(pdf_bytes))],
        "extractedTexts": []
    }


def _extract_office(name: str, data: bytes) -> dict:
    suffix = pathlib.Path(name).suffix.lower()
    if suffix not in {".doc", ".docx", ".ppt", ".pptx"}:
        raise HTTPException(400, "supported extensions: .doc/.docx/.ppt/.pptx")

    with TemporaryDirectory() as td:
        in_path = pathlib.Path(td) / f"input{suffix}"
        in_path.write_bytes(data)

        # legacy → OpenXML so the rest stays identical
        if suffix in {".doc", ".ppt"}:
            target_ext = "docx" if suffix == ".doc" else "pptx"
            in_path = soffice_convert(in_path, in_path.parent, target_fmt=target_ext)
            suffix = f".{target_ext}"

        files: List[dict]    = []
        img_hash_by_member   : Dict[str, str] = {}
        seen_hashes: Set[str] = set()

        with open_office_zip(in_path) as zf:
            media_members = [n for n in zf.namelist() if "/media/" in n]
            for member in media_members:
                data = zf.read(member)
                h    = sha1(data)

                # record every path → hash so placeholders always resolve
                img_hash_by_member[member] = h
                img_hash_by_member[pathlib.Path(member).name] = h

                if h in seen_hashes:
                    continue  # duplicate payload, skip in output list
                seen_hashes.add(h)

                ext  = pathlib.Path(member).suffix.lstrip(".") or "png"
                mime = mimetypes.guess_type(member)[0] or f"image/{ext}"
                files.append(file_json(f"img_{len(files)+1:04}.{ext}", mime, data, hash_hex=h))

            try:
                text = (extract_docx_text_with_placeholders if suffix == ".docx" else extract_pptx_text_with_placeholders)(zf, img_hash_by_member)
            except KeyError as e:  # mandatory part missing, e.g. word/document.xml
                raise HTTPException(status.HTTP_422_UNPROCESSABLE_ENTITY, f"invalid document: {e}")

    return {
        "files": files,
        "extractedTexts": [{"name": pathlib.Path(name).stem, "text": text}]
    }


def _extract_pdf(name: str, data: bytes) -> dict:
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception:
        raise HTTPException(status.HTTP_422_UNPROCESSABLE_ENTITY, "file is not a valid PDF")

    with doc:
        unique_by_hash, hash_by_xref = collect_pdf_images(doc)
        text = extract_pdf_text_with_placeholders(doc, hash_by_xref, unique_by_hash)

    files = [
        file_json(
            f"img_{i:04}.{img['ext'] or 'png'}",
            "image/jpeg" if img["ext"] == "jpg" else f"image/{img['ext']}",
            img["image"],
            hash_hex=h,
        )
        for i, (h, img) in enumerate(unique_by_hash.items(), 1)
    ]

    return {
        "files": files,
        "extractedTexts": [{"name": pathlib.Path(name).stem, "text": text}]
    }

# --------------------------------------------------------------------------- #
# Routes                                                                       #
# --------------------------------------------------------------------------- #

@app.post("/convert/docx-to-pdf")
async def docx_to_pdf(file: UploadFile = File(...)):
    """Convert DOCX → PDF."""
    name, data = await read_upload(file)
    validate_suffix(name, ".docx")
    return JSONResponse(await run_in_threadpool(_convert_to_pdf, name, data, ".docx"))


@app.post("/convert/pptx-to-pdf")
async def pptx_to_pdf(file: UploadFile = File(...)):
    """Convert PPTX → PDF."""
    name, data = await read_upload(file)
    validate_suffix(name, ".pptx")
    return JSONResponse(await run_in_threadpool(_convert_to_pdf, name, data, ".pptx"))


@app.post("/extract/images-office")
async def extract_images_office(file: UploadFile = File(...)):
    """Extract deduplicated images + placeholder‑rich text from DOCX/PPTX (+ legacy DOC/PPT)."""
    name, data = await read_upload(file)
    return JSONResponse(await run_in_threadpool(_extract_office, name, data))


@app.post("/extract/images")
async def extract_images(file: UploadFile = File(...)):
    """Extract deduplicated images + placeholder‑rich text from a PDF."""
    name, data = await read_upload(file)
    validate_suffix(name, ".pdf")
    return JSONResponse(await run_in_threadpool(_extract_pdf, name, data))


# --------------------------------------------------------------------------- #
# Health probe                                                                 #
# --------------------------------------------------------------------------- #

@app.get("/healthz", include_in_schema=False)
async def healthz():
    return {"status": "ok"}
