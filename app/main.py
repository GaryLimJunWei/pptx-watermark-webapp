from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, FileResponse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import io
import zipfile
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parents[1]
STATIC_DIR = APP_ROOT / "static"

app = FastAPI()


def mm_to_inches(mm: float) -> float:
    return mm / 25.4


def add_name_to_all_slides(pptx_bytes: bytes, name: str) -> bytes:
    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
            if "[Content_Types].xml" not in zf.namelist():
                raise ValueError("Not a valid Office file.")
    except Exception:
        raise ValueError("Uploaded file is not a valid .pptx")

    prs = Presentation(io.BytesIO(pptx_bytes))

    # ---- Watermark Position (edit once if needed) ----
    margin_right_mm = 12
    margin_bottom_mm = 10
    box_width_mm = 70
    box_height_mm = 10

    font_size_pt = 12
    font_rgb = RGBColor(80, 80, 80)

    tag_shape_name = "__WATERMARK_NAME__"

    for slide in prs.slides:
        # Remove old watermark if exists
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

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue()


@app.get("/")
def home():
    return FileResponse(STATIC_DIR / "index.html")


@app.post("/process")
async def process_pptx(file: UploadFile = File(...), name: str = Form(...)):

    name = (name or "").strip()
    if not name:
        raise HTTPException(status_code=400, detail="Name is required.")

    if not (file.filename or "").lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Upload a .pptx file.")

    raw = await file.read()

    if len(raw) > 50 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File too large (max 50 MB).")

    try:
        processed = add_name_to_all_slides(raw, name)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))

    output_stream = io.BytesIO(processed)
    output_stream.seek(0)

    base = (file.filename or "presentation.pptx").rsplit(".pptx", 1)[0]
    download_name = f"{base}__named.pptx"

    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{download_name}"'},
    )
