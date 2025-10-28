from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
import zipfile
import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
import re
import logging
import io
from urllib.parse import urlparse

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def get_value_for_field(row, field):
    try:
        if field in row and not pd.isna(row[field]):
            return str(row[field]).strip()
        return ""
    except: return ""

def is_image_path(val):
    return val and isinstance(val, str) and bool(re.search(r'\.(jpg|jpeg|png|gif|bmp)$', val.lower()))

def find_image_path(value, images_dir):
    try:
        clean_value = value.replace("images/", "", 1) if value.startswith("images/") else value
        candidate = os.path.join(images_dir, clean_value)
        if os.path.exists(candidate):
            return candidate
        for root, _, files in os.walk(images_dir):
            for file in files:
                if file.lower() == clean_value.lower():
                    return os.path.join(root, file)
        return None
    except: return None

# ----------------------------------------------------------------------
# TEXT REPLACEMENT (DASHES + LINKS)
# ----------------------------------------------------------------------
def replace_text_in_obj(obj, row):
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field in matches:
                        val = get_value_for_field(row, field)
                        if not is_image_path(val):
                            # DASH FIX
                            escaped = re.escape(f"{{{{{field}}}}}")
                            flexible = (
                                escaped
                                .replace('\\{\\{', r'\s*\{\{\s*')
                                .replace('\\}\\}', r'\s*\}\}')
                                .replace('\\-', r'[-–—]')
                                .replace('\\ ', r'\s*')
                            )
                            match = re.search(flexible, run.text, re.UNICODE)
                            if match:
                                placeholder_text = match.group(0)
                                # LINK STYLING
                                if field.lower().endswith("link"):
                                    try:
                                        result = urlparse(val)
                                        if all([result.scheme, result.netloc]):
                                            run.text = run.text.replace(placeholder_text, val, 1)
                                            run.font.color.rgb = RGBColor(0, 0, 255)
                                            run.font.underline = True
                                            continue
                                    except: pass
                                # NORMAL TEXT
                                run.text = run.text.replace(placeholder_text, val, 1)
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

# ----------------------------------------------------------------------
# IMAGE REPLACEMENT (unchanged)
# ----------------------------------------------------------------------
def replace_images_on_shape(shape, row, images_dir):
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
            matches = placeholder_pattern.findall(full_text)
            for field in matches:
                val = get_value_for_field(row, field)
                if is_image_path(val):
                    img_path = find_image_path(val, images_dir)
                    if img_path:
                        for p in shape.text_frame.paragraphs:
                            for r in p.runs:
                                r.text = r.text.replace(f"{{{{{field}}}}}", "")
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        sp = shape._element
                        sp.getparent().remove(sp)
                        slide = shape.part.slide
                        slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                        return
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

# ----------------------------------------------------------------------
# MAIN ENDPOINT (TABLES ADDED SAFELY)
# ----------------------------------------------------------------------
@app.post("/api/generate")
async def generate(excel: UploadFile = File(...), ppt: UploadFile = File(...), images: UploadFile = File(None)):
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            excel_path = os.path.join(tmpdir, excel.filename or "content.xlsx")
            ppt_path = os.path.join(tmpdir, ppt.filename or "template.pptx")
            images_dir = os.path.join(tmpdir, "images")

            with open(excel_path, "wb") as f: f.write(await excel.read())
            with open(ppt_path, "wb") as f: f.write(await ppt.read())

            if images:
                zip_path = os.path.join(tmpdir, images.filename or "images.zip")
                with open(zip_path, "wb") as f: f.write(await images.read())
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(images_dir)
                logger.info(f"Extracted images to: {images_dir}")

            df = pd.read_excel(excel_path)
            df.columns = [col.strip() for col in df.columns]
            prs = Presentation(ppt_path)

            # PASS 1: Images
            logger.info("Processing slides – Step 1: Images")
            for i, row in df.iterrows():
                if i >= len(prs.slides): break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_images_on_shape(shape, row, images_dir if images else tmpdir)
                    # TABLES: SAFE
                    if hasattr(shape, "table"):
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_images_on_shape(cell, row, images_dir if images else tmpdir)

            # PASS 2: Text
            logger.info("Processing slides – Step 2: Text")
            for i, row in df.iterrows():
                if i >= len(prs.slides): break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_text_in_obj(shape, row)
                    # TABLES: SAFE
                    if hasattr(shape, "table"):
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_text_in_obj(cell, row)

            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            prs.save(output_file)

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
