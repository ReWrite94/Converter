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
      "…plain text of processed document with {{<<<sha1>>>}} placeholders…"
  ]
}

• Every distinct image is returned only once (duplicate detection via SHA‑1).
• In the text, each image position is replaced by its hash placeholder.
"""
from __future__ import annotations

import base64
import hashlib
import html
import mimetypes
import pathlib
import re
import subprocess
import zipfile
from tempfile import TemporaryDirectory
from typing import Dict, List, Set, Tuple

import fitz  # PyMuPDF

PDF_TEXT_BLOCK  = getattr(fitz, "PDF_TEXT_BLOCK", 0)
PDF_IMAGE_BLOCK = getattr(fitz, "PDF_IMAGE_BLOCK", 1)

from fastapi import Depends, FastAPI, File, HTTPException, UploadFile, status
from fastapi.responses import JSONResponse

from security import api_key_auth

app = FastAPI(
    title="Doc ⇄ PDF Service (JSON)",
    version="2.4.6",  # sentinel placeholders keep 40‑hex hashes intact
    dependencies=[Depends(api_key_auth)],
)

# --------------------------------------------------------------------------- #
# Helper functions                                                             #
# --------------------------------------------------------------------------- #

def to_base64(data: bytes) -> str:
    """bytes → base64‑string (utf‑8)"""
    return base64.b64encode(data).decode()


def sha1(data: bytes) -> str:
    """Return 40‑char SHA‑1 hex digest of *data*."""
    return hashlib.sha1(data).hexdigest()


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


# --------------------------------------------------------------------------- #
# LibreOffice conversion                                                       #
# --------------------------------------------------------------------------- #

def soffice_convert(src: pathlib.Path, outdir: pathlib.Path, *, target_fmt: str) -> pathlib.Path:
    """Convert *src* to *target_fmt* (e.g. pdf/docx) using LibreOffice."""
    cmd = [
        "soffice",
        "--headless",
        "--convert-to",
        target_fmt,
        "--outdir",
        str(outdir),
        str(src),
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"LibreOffice error ({res.returncode}):\n{res.stderr or res.stdout}")
    return outdir / src.with_suffix(f".{target_fmt}").name


# --------------------------------------------------------------------------- #
# Validation helper                                                            #
# --------------------------------------------------------------------------- #

def validate_suffix(filename: str, *allowed: str) -> None:
    if not filename.lower().endswith(allowed):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"expected file extension {'/'.join(allowed)}",
        )

# --------------------------------------------------------------------------- #
# Text extraction helpers                                                      #
# --------------------------------------------------------------------------- #

PDF_TEXT_BLOCK = 0
PDF_IMAGE_BLOCK = 1

_SENTINEL = "|||IMG"  # protects hashes while stripping XML tags

_RE_REL  = re.compile(r"Id=\"([^\"]+)\"[^>]*Target=\"([^\"]+)\"")
_RE_BLIP = re.compile(r"<a:blip[^>]+?r:embed=\"([^\"]+)\"[^>]*/>")
_RE_TAGS = re.compile(r"<[^>]+>")


def _strip_xml_keep_sentinels(xml: str) -> str:
    no_tags = _RE_TAGS.sub(" ", xml)
    return re.sub(r"\s+", " ", no_tags).strip()


def _sentinel_to_placeholder(text: str) -> str:
    pattern = re.compile(r"\|\|\|IMG([0-9a-f]{40})\|\|\|")
    return pattern.sub(lambda m: "{{<<<" + m.group(1) + ">>>}}", text)

# Schlüssel kann int (xref) oder Tuple[int,int] (page_idx, inline_number) sein
ImageKey = Tuple[int, int] | int

def build_image_hashes(doc: fitz.Document) -> Dict[ImageKey, str]:
    image_hashes: Dict[ImageKey, str] = {}

    for page_idx, page in enumerate(doc):
        # 1) richtige Bilder mit xref
        for img in page.get_images(full=True):          # (xref, smask, ...)
            xref = img[0]
            img_bytes = doc.extract_image(xref)['image']
            h = hashlib.sha1(img_bytes).hexdigest()
            image_hashes[xref] = h

        # 2) inline-Bilder (PyMuPDF ≥1.21 liefert sie als eigene Blöcke)
        for blk in page.get_text("dict")["blocks"]:
            if blk["type"] == fitz.PDF_IMAGE_BLOCK and blk.get("xref") is None:
                inline_bytes = blk["image"]
                h = hashlib.sha1(inline_bytes).hexdigest()
                key = (page_idx, blk["number"])          # eindeutig im Dokument!
                image_hashes[key] = h
    return image_hashes

def _text_from_rawdict_block(blk: dict) -> str:
    """
    Liefert den reinen Textinhalt eines RAWDICT- oder DICT-Textblocks.
    Kompatibel mit allen PyMuPDF-Versionen (0.x – 1.x).
    """
    # 1) Neuere DICT-Outputs (PyMuPDF ≥ 1.19) haben 'text'
    if blk.get("text"):
        return blk["text"]

    # 2) RAWDICT / DICT: Hierarchien  blocks → lines → spans
    if "lines" in blk:
        pieces: list[str] = []
        for line in blk["lines"]:
            for span in line.get("spans", []):
                if span.get("text"):                     # DICT
                    pieces.append(span["text"])
                elif span.get("chars"):                  # RAWDICT ≥ 1.21
                    pieces.append("".join(ch.get("c", "") for ch in span["chars"]))
            pieces.append("\n")                          # Zeilenumbruch beibehalten
        return "".join(pieces).rstrip("\n")

    # 3) Sehr alte RAWDICT-Varianten (< 1.18) benutzen 'items'
    if "items" in blk:
        return "".join(
            it.get("text", it.get("c", ""))
            for it in blk["items"]
        )

    # Fallback
    return ""

# --------------------------- PDF ------------------------------------------- #

def extract_pdf_text_with_placeholders(
    doc: fitz.Document,
    image_hashes: Optional[Dict[ImageKey, str]] = None
) -> str:
    """
    PDF-Text mit Platzhaltern extrahieren.

        • Große Bilder (≥48 px, ≥4 000 px², ≥1 KB)
          → {{<<<sha1>>>}}
        • Kleinere Bilder/Icon
          → {{<<<icon>>>}}
    """
    # Schwellwerte (identisch zur Sammel-Schleife in /extract/images)
    MIN_W, MIN_H = 48, 48
    MIN_AREA     = 4_000
    MIN_BYTES    = 1_024

    if image_hashes is None:
        image_hashes = {}

    parts: List[str] = []

    for page_idx, page in enumerate(doc):
        blocks = page.get_text("rawdict")["blocks"]

        for blk in blocks:
            if blk["type"] == PDF_TEXT_BLOCK:
                parts.append(_text_from_rawdict_block(blk))

            elif blk["type"] == PDF_IMAGE_BLOCK:
                # Schlüssel bestimmen
                if blk.get("xref") is not None:
                    key: ImageKey = blk["xref"]
                    get_bytes = lambda: doc.extract_image(key)["image"]
                    get_dims  = lambda: (
                        doc.extract_image(key)["width"],
                        doc.extract_image(key)["height"],
                    )
                else:  # Inline-Bild
                    key = (page_idx, blk["number"])
                    get_bytes = lambda: blk["image"]
                    get_dims  = lambda: (blk["width"], blk["height"])

                # Bild bereits akzeptiert → Platzhalter mit Hash
                if key in image_hashes:
                    parts.append(f"{{<<<{image_hashes[key]}>>>}}")
                    continue

                # Neu bewerten (nur für Inline oder selten übersehene XObjects)
                img_bytes = get_bytes()
                w, h      = get_dims()

                if (
                    w >= MIN_W and h >= MIN_H and
                    w * h >= MIN_AREA and
                    len(img_bytes) >= MIN_BYTES
                ):
                    # Groß genug → Hash berechnen, merken, Platzhalter ausgeben
                    hsh = hashlib.sha1(img_bytes).hexdigest()
                    image_hashes[key] = hsh
                    parts.append(f"{{<<<{hsh}>>>}}")
                else:
                    # Zu klein → Icon-Marker
                    parts.append("{{<<<icon>>>}}")

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

def extract_pptx_text_with_placeholders(zf: zipfile.ZipFile, img_hash_by_member: Dict[str, str]) -> str:
    slide_paths = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
    chunks: List[str] = []

    for slide_xml_path in sorted(slide_paths):
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
# Routes                                                                       #
# --------------------------------------------------------------------------- #

@app.post("/convert/docx-to-pdf")
async def docx_to_pdf(file: UploadFile = File(...)):
    """Convert DOCX → PDF."""
    validate_suffix(file.filename, ".docx")
    with TemporaryDirectory() as td:
        src = pathlib.Path(td) / file.filename
        src.write_bytes(await file.read())
        pdf_path = soffice_convert(src, src.parent, target_fmt="pdf")
        pdf_bytes = pdf_path.read_bytes()
    return JSONResponse({
        "files": [file_json(pdf_path.name, "application/pdf", pdf_bytes, hash_hex=sha1(pdf_bytes))],
        "extractedTexts": []
    })


@app.post("/convert/pptx-to-pdf")
async def pptx_to_pdf(file: UploadFile = File(...)):
    """Convert PPTX → PDF."""
    validate_suffix(file.filename, ".pptx")
    with TemporaryDirectory() as td:
        src = pathlib.Path(td) / file.filename
        src.write_bytes(await file.read())
        pdf_path = soffice_convert(src, src.parent, target_fmt="pdf")
        pdf_bytes = pdf_path.read_bytes()
    return JSONResponse({
        "files": [file_json(pdf_path.name, "application/pdf", pdf_bytes, hash_hex=sha1(pdf_bytes))],
        "extractedTexts": []
    })


@app.post("/extract/images-office")
async def extract_images_office(file: UploadFile = File(...)):
    """Extract deduplicated images + placeholder‑rich text from DOCX/PPTX (+ legacy DOC/PPT)."""
    name   = file.filename
    suffix = pathlib.Path(name).suffix.lower()

    with TemporaryDirectory() as td:
        in_path = pathlib.Path(td) / name
        in_path.write_bytes(await file.read())

        # legacy → OpenXML so the rest stays identical
        if suffix in {".doc", ".ppt"}:
            target_ext = "docx" if suffix == ".doc" else "pptx"
            in_path = soffice_convert(in_path, in_path.parent, target_fmt=target_ext)
            suffix = f".{target_ext}"
        elif suffix not in {".docx", ".pptx"}:
            raise HTTPException(400, "supported extensions: .doc/.docx/.ppt/.pptx")

        files: List[dict]    = []
        img_hash_by_member   : Dict[str, str] = {}
        seen_hashes: Set[str] = set()

        with zipfile.ZipFile(in_path) as zf:
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

            text = (extract_docx_text_with_placeholders if suffix == ".docx" else extract_pptx_text_with_placeholders)(zf, img_hash_by_member)

    return JSONResponse({
        "files": files,
        "extractedTexts": [{"name": pathlib.Path(file.filename).stem, "text": text}]
    })


@app.post("/extract/images")
async def extract_images(file: UploadFile = File(...)):
    """Extract deduplicated images + placeholder‑rich text from a PDF."""
    validate_suffix(file.filename, ".pdf")

    MIN_W, MIN_H = 48, 48
    MIN_AREA     = 4_000
    MIN_BYTES    = 1_024

    unique_by_hash: Dict[str, dict] = {}
    hash_by_xref : Dict[int, str]   = {}

    with TemporaryDirectory() as td:
        pdf_path = pathlib.Path(td) / file.filename
        pdf_path.write_bytes(await file.read())
        doc = fitz.open(pdf_path)

        for page in doc:
            for xref, *_ in page.get_images(full=True):
                info = doc.extract_image(xref)
                w, h = info["width"], info["height"]

                if info["cs-name"] == "DeviceGray" and info["bpc"] == 1:
                    continue  # 1‑bit mask images
                if w < MIN_W or h < MIN_H or w * h < MIN_AREA:
                    continue
                if len(info["image"]) < MIN_BYTES:
                    continue

                hsh = sha1(info["image"])
                hash_by_xref[xref] = hsh
                if hsh in unique_by_hash:
                    continue
                unique_by_hash[hsh] = info

        text = extract_pdf_text_with_placeholders(doc, hash_by_xref)

    files = [
        file_json(
            f"img_{i:04}.{img['ext'] or 'png'}",
            "image/jpeg" if img["ext"] == "jpg" else f"image/{img['ext']}",
            img["image"],
            hash_hex=h,
        )
        for i, (h, img) in enumerate(unique_by_hash.items(), 1)
    ]

    return JSONResponse({
        "files": files,
        "extractedTexts": [{"name": pathlib.Path(file.filename).stem, "text": text}]
    })


# --------------------------------------------------------------------------- #
# Health probe                                                                 #
# --------------------------------------------------------------------------- #

@app.get("/healthz", include_in_schema=False)
async def healthz():
    return {"status": "ok"}
