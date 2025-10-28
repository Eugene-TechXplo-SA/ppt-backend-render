from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
import zipfile
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import re
import logging
import io
from urllib.parse import urlparse

# ----------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For testing
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ----------------------------------------------------------------------
# Helper utilities
# ----------------------------------------------------------------------
def get_value_for_field(row, field):
    try:
        if field in row and not pd.isna(row[field]):
            val = str(row[field]).strip()
            logger.debug(f"Got value for {field}: {val}")
            return val
        logger.warning(f"No value for {field}")
        return ""
    except Exception as e:
        logger.error(f"Error in get_value_for_field for {field}: {str(e)}")
        return ""

def is_image_path(val):
    if val and isinstance(val, str):
        result = bool(re.search(r'\.(jpg|jpeg|png|gif|bmp)$', val.lower()))
        logger.debug(f"Checking if {val} is image: {result}")
        return result
    return False

def find_image_path(value, images_dir):
    try:
        if not value or not images_dir:
            logger.warning(f"No value or images_dir for {value}")
            return None
        clean_value = value.replace("images/", "", 1) if value.startswith("images/") else value
        candidate = os.path.join(images_dir, clean_value)
        logger.debug(f"Looking for image at: {candidate}")
        if os.path.exists(candidate):
            logger.info(f"Found image: {candidate}")
            return candidate
        # case-insensitive fallback
        for root, _, files in os.walk(images_dir):
            for file in files:
                if file.lower() == clean_value.lower():
                    full_path = os.path.join(root, file)
                    logger.info(f"Found case-insensitive match: {full_path}")
                    return full_path
        logger.warning(f"Image not found: {candidate}")
        return None
    except Exception as e:
        logger.error(f"Error in find_image_path for {value}: {str(e)}")
        return None

# ----------------------------------------------------------------------
# RECURSIVE TEXT-FRAME WALKER
# ----------------------------------------------------------------------
def _iter_text_frames(shape):
    """Yield every text_frame that belongs to *shape* (groups, placeholders, tables)."""
    if hasattr(shape, "text_frame") and shape.text_frame is not None:
        yield shape.text_frame

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            yield from _iter_text_frames(child)

    if hasattr(shape, "is_placeholder") and shape.is_placeholder:
        if hasattr(shape, "text_frame") and shape.text_frame is not None:
            yield shape.text_frame

# ----------------------------------------------------------------------
# BULLET-PROOF REPLACEMENT (handles en-dashes, spaces, punctuation)
# ----------------------------------------------------------------------
def _replace_in_text_frame(tf, row, is_image_pass=False):
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")

    for paragraph in tf.paragraphs:
        for run in paragraph.runs:
            original_text = run.text
            if not original_text:
                continue

            matches = placeholder_pattern.findall(original_text)
            if not matches:
                continue

            new_text = original_text

            for field_raw in matches:
                # ---- CLEAN the field name -------------------------------------------------
                field_clean = field_raw.strip()
                field_clean = re.sub(r'\s+', ' ', field_clean)          # collapse spaces
                field_clean = field_clean.replace('–', '-').replace('—', '-')  # normalise dashes

                # ---- Find Excel column ----------------------------------------------------
                col = next((c for c in row.index if c.lower().strip() == field_clean.lower()), None)
                if not col:
                    logger.debug(f"Column not found for cleaned field: {field_clean}")
                    continue
                val = get_value_for_field(row, col)

                # ---- Build a flexible regex that finds the *exact* placeholder in the original text
                escaped = re.escape(f"{{{{{field_raw}}}}}")
                flexible = (
                    escaped
                    .replace('\\{\\{', r'\s*\{\{\s*')      # spaces inside {{
                    .replace('\\}\\}', r'\s*\}\}')        # spaces inside }}
                    .replace('\\-', r'[-–—]')             # any dash
                    .replace('\\ ', r'\s*')               # any whitespace
                )
                match = re.search(flexible, new_text, re.UNICODE)
                if not match:
                    logger.warning(f"Placeholder not located in run: {field_raw}")
                    continue

                placeholder_text = match.group(0)

                # ---- IMAGE PASS -----------------------------------------------------------
                if is_image_pass and is_image_path(val):
                    new_text = new_text.replace(placeholder_text, "", 1)
                    continue

                # ---- TEXT PASS ------------------------------------------------------------
                if col.lower().endswith("link"):
                    try:
                        result = urlparse(val)
                        if all([result.scheme, result.netloc]):
                            new_text = new_text.replace(placeholder_text, val, 1)
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
                            logger.info(f"Replaced {field_raw} with hyperlink: {val}")
                        else:
                            new_text = new_text.replace(placeholder_text, val, 1)
                    except Exception:
                        new_text = new_text.replace(placeholder_text, val, 1)
                else:
                    new_text = new_text.replace(placeholder_text, val, 1)
                    logger.info(f"Replaced {field_raw} with text: {val}")

            run.text = new_text

# ----------------------------------------------------------------------
# IMAGE REPLACEMENT (uses the new helper)
# ----------------------------------------------------------------------
def replace_images_on_shape(shape, row, images_dir):
    """First pass – only images (recursive)."""
    for tf in _iter_text_frames(shape):
        _replace_in_text_frame(tf, row, is_image_pass=True)

    # ---- actual picture insertion (unchanged) ---------------------------------
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    if hasattr(shape, "has_text_frame") and shape.has_text_frame:
        full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
        matches = placeholder_pattern.findall(full_text)
        for field in matches:
            val = get_value_for_field(row, field)
            if is_image_path(val):
                img_path = find_image_path(val, images_dir)
                if img_path:
                    logger.info(f"Inserting image for {field} from {img_path}")
                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            r.text = r.text.replace(f"{{{{{field}}}}}", "")
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    sp = shape._element
                    sp.getparent().remove(sp)
                    slide = shape.part.slide
                    slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                    logger.info(f"Successfully inserted image: {img_path}")
                    return

# ----------------------------------------------------------------------
# TEXT REPLACEMENT (uses the new helper)
# ----------------------------------------------------------------------
def replace_text_in_obj(obj, row):
    """Second pass – plain text & hyperlinks (recursive)."""
    for tf in _iter_text_frames(obj):
        _replace_in_text_frame(tf, row, is_image_pass=False)

# ----------------------------------------------------------------------
# MAIN ENDPOINT (unchanged except the two-pass calls)
# ----------------------------------------------------------------------
@app.post("/api/generate")
async def generate(excel: UploadFile = File(...), ppt: UploadFile = File(...), images: UploadFile = File(None)):
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            excel_filename = excel.filename or "content.xlsx"
            ppt_filename = ppt.filename or "template_client.pptx"
            excel_path = os.path.join(tmpdir, excel_filename)
            ppt_path = os.path.join(tmpdir, ppt_filename)
            images_dir = os.path.join(tmpdir, "images")
            logger.info(f"Saving files: excel={excel_path}, ppt={ppt_path}")

            # ----- read excel -----
            excel_content = await excel.read()
            if not excel_content:
                raise HTTPException(status_code=400, detail="Excel file is empty")
            with open(excel_path, "wb") as f:
                f.write(excel_content)

            # ----- read ppt -----
            ppt_content = await ppt.read()
            if not ppt_content:
                raise HTTPException(status_code=400, detail="PowerPoint file is empty")
            with open(ppt_path, "wb") as f:
                f.write(ppt_content)

            # ----- optional images zip -----
            if images:
                zip_filename = images.filename or "images.zip"
                zip_path = os.path.join(tmpdir, zip_filename)
                zip_content = await images.read()
                if not zip_content:
                    raise HTTPException(status_code=400, detail="Images ZIP file is empty")
                with open(zip_path, "wb") as f:
                    f.write(zip_content)
                try:
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(images_dir)
                    logger.info(f"Extracted images to: {images_dir}")
                except zipfile.BadZipFile as e:
                    logger.error(f"Invalid ZIP file: {str(e)}")
                    raise HTTPException(status_code=400, detail=f"Invalid ZIP file: {str(e)}")

            # ----- load data -----
            df = pd.read_excel(excel_path)
            df.columns = [col.strip() for col in df.columns]
            prs = Presentation(ppt_path)

            # ----- IMAGE PASS -----
            logger.info("Processing slides – Step 1: Images")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_images_on_shape(shape, row, images_dir if images else tmpdir)
                    if shape.shape_type == 19:  # TABLE
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_images_on_shape(cell, row, images_dir if images else tmpdir)

            # ----- TEXT PASS -----
            logger.info("Processing slides – Step 2: Text")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_text_in_obj(shape, row)
                    if shape.shape_type == 19:  # TABLE
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_text_in_obj(cell, row)

            # ----- save output -----
            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            prs.save(output_file)

            # ----- stream back -----
            with open(output_file, "rb") as f:
                file_content = f.read()
            return StreamingResponse(
                io.BytesIO(file_content),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": "attachment; filename=Client_Presentation.pptx"}
            )

        except Exception as e:
            logger.error(f"Error in /api/generate: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")
