from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import io
import os
import re
import json
import tempfile
import subprocess
import zipfile
import smtplib
from email.message import EmailMessage
from pathlib import Path


APP_ROOT = Path(__file__).resolve().parents[1]
STATIC_DIR = APP_ROOT / "static"

# --- Google Drive folder id (from your link) ---
DRIVE_FOLDER_ID = "1YBK7WC2pbS49sBU7hZyyehXvyQzhwkud"

# --- Email notification target ---
NOTIFY_TO = "garylimjunwei@gmail.com"

# --- SMTP settings (set in Render env vars) ---
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")          # e.g. garylimjunwei@gmail.com
SMTP_APP_PASSWORD = os.environ.get("SMTP_APP_PASSWORD", "")  # Gmail App Password (16 chars)

# --- Google Service Account JSON (set in Render env var) ---
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

app = FastAPI()


def mm_to_inches(mm: float) -> float:
    return mm / 25.4


def validate_pptx_bytes(pptx_bytes: bytes) -> None:
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
            if "[Content_Types].xml" not in zf.namelist():
                raise ValueError("Not a valid Office file.")
    except Exception:
        raise ValueError("Uploaded file is not a valid .pptx")


def sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^a-zA-Z0-9._ -]+", "_", name)
    return name[:180] or "upload.pptx"


def get_drive_service():
    if not GOOGLE_SERVICE_ACCOUNT_JSON.strip():
        raise RuntimeError("Missing GOOGLE_SERVICE_ACCOUNT_JSON env var.")

    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    return build("drive", "v3", credentials=creds)


def upload_original_to_drive(original_bytes: bytes, original_filename: str) -> str:
    """
    Uploads ONLY the original uploaded PPTX into the target Drive folder.
    Returns fileId.
    """
    service = get_drive_service()

    fh = io.BytesIO(original_bytes)
    media = MediaIoBaseUpload(
        fh,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=False,
    )

    file_metadata = {
        "name": original_filename,
        "parents": [DRIVE_FOLDER_ID],
    }

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id",
        supportsAllDrives=True,
    ).execute()

    return created["id"]


def send_notification_email(original_filename: str, drive_file_id: str) -> None:
    """
    Sends you a simple email: new upload + filename + Drive file id.
    Uses Gmail SMTP with App Password (recommended).
    Gmail requires 2-step verification + app password. :contentReference[oaicite:2]{index=2}
    """
    if not (SMTP_USER and SMTP_APP_PASSWORD):
        # If you haven't set SMTP env vars, skip sending rather than failing the user download.
        return

    msg = EmailMessage()
    msg["Subject"] = "New PowerPoint upload received"
    msg["From"] = SMTP_USER
    msg["To"] = NOTIFY_TO
    msg.set_content(
        f"A new PowerPoint was uploaded.\n\n"
        f"File: {original_filename}\n"
        f"Drive fileId: {drive_file_id}\n"
        f"(Uploaded original only)\n"
    )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_APP_PASSWORD)
        server.send_message(msg)


def add_name_to_all_slides(pptx_bytes: bytes, name: str) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Watermark placement: bottom-right (consistent)
    margin_right_mm = 12
    margin_bottom_mm = 10
    box_width_mm = 70
    box_height_mm = 10

    font_size_pt = 12
    font_rgb = RGBColor(255, 0, 0)  # RED

    tag_shape_name = "__WATERMARK_NAME__"

    for slide in prs.slides:
        # remove prior watermark if present
        to_remove = []
        for shape in slide.shapes:
            if getattr(shape, "name", "") == tag_shape_name:
                to_remove.append(shape)
        for shape in to_remove:
            el = shape._element
            el.getparent().remove(el)

        sw = prs.slide_width
        sh = prs.slide_height

        width = Inches(mm_to_inches(box_width_mm))
        height = Inches(mm_to_inches(box_height_mm))

        left = sw - width - Inches(mm_to_inches(margin_right_mm))
        top = sh - height - Inches(mm_to_inches(margin_bottom_mm))

        textbox = slide.shapes.add_textbox(left, top, width, height)
        try:
            textbox.name = tag_shape_name
        except Exception:
            pass

        tf = textbox.text_frame
        tf.clear()

        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = name

        p.alignment = PP_ALIGN.RIGHT
        run.font.size = Pt(font_size_pt)
        run.font.color.rgb = font_rgb

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def convert_pptx_to_pdf(pptx_path: str, out_dir: str) -> str:
    """
    Uses LibreOffice headless to convert PPTX -> PDF.
    Command form is standard: soffice --headless --convert-to pdf --outdir ...
    :contentReference[oaicite:3]{index=3}
    """
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        pptx_path,
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"PDF conversion failed: {proc.stderr[-800:]}")

    base = Path(pptx_path).stem
    pdf_path = str(Path(out_dir) / f"{base}.pdf")
    if not Path(pdf_path).exists():
        # LibreOffice sometimes outputs with same stem; if not found, scan for a pdf
        pdfs = list(Path(out_dir).glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("PDF conversion produced no output.")
        pdf_path = str(pdfs[0])
    return pdf_path


@app.get("/")
def home():
    return FileResponse(STATIC_DIR / "index.html")


@app.post("/process")
async def process(file: UploadFile = File(...), name: str = Form(...)):
    name = (name or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="Name is required.")

    if not (file.filename or "").lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Upload a .pptx file.")

    raw = await file.read()
    if len(raw) > 50 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File too large (max 50 MB).")

    try:
        validate_pptx_bytes(raw)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    original_filename = sanitize_filename(file.filename or "upload.pptx")

    # 1) Upload ORIGINAL to Drive + send email (silent to user)
    drive_file_id = ""
    try:
        drive_file_id = upload_original_to_drive(raw, original_filename)
        send_notification_email(original_filename, drive_file_id)
    except Exception:
        # Don’t block the user download if Drive/email fails
        pass

    # 2) Add red name, then convert to PDF
    try:
        watermarked = add_name_to_all_slides(raw, name)
        with tempfile.TemporaryDirectory() as td:
            in_pptx = Path(td) / "watermarked.pptx"
            in_pptx.write_bytes(watermarked)

            pdf_path = convert_pptx_to_pdf(str(in_pptx), td)
            # Return PDF
            download_name = Path(original_filename).stem + "__named.pdf"
            return FileResponse(
                pdf_path,
                media_type="application/pdf",
                filename=download_name,
            )
    except RuntimeError as e:
        raise HTTPException(status_code=500, detail=str(e))
    except Exception:
        raise HTTPException(status_code=500, detail="Processing failed.")
