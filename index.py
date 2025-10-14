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

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For testing
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_value_for_field(row, field):
    """Return a safe string value for a field from Excel—trim spaces."""
    try:
        if field in row and not pd.isna(row[field]):
            val = str(row[field]).strip()  # Trim spaces for "Site Number "
            return val
        return ""
    except Exception as e:
        logger.error(f"Error in get_value_for_field for field {field}: {str(e)}")
        return ""

def find_image_path(value, images_dir):
    """Try to resolve image path in images_dir—strip 'images/' if present."""
    try:
        if not value or not images_dir:
            return None
        # Strip 'images/' prefix if in value (e.g., "images/pic.jpg" -> "pic.jpg")
        clean_value = value.replace("images/", "", 1) if value.startswith("images/") else value
        candidate = os.path.join(images_dir, clean_value)
        if os.path.exists(candidate):
            return candidate
        logger.warning(f"Image not found: {candidate}")
        return None
    except Exception as e:
        logger.error(f"Error in find_image_path for {value}: {str(e)}")
        return None

def replace_text_in_obj(obj, row):
    """Replace text placeholders inside shape or cell."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field in matches:
                        val = get_value_for_field(row, field)
                        run.text = run.text.replace(f"{{{{{field}}}}}", val)
                        if field.lower() == "link" and val:
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

def replace_images_on_shape(shape, row, images_dir):
    """Replace placeholders with images if the value points to a file."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
            matches = placeholder_pattern.findall(full_text)
            for field in matches:
                img_path = find_image_path(get_value_for_field(row, field), images_dir)
                if img_path:
                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            r.text = r.text.replace(f"{{{{{field}}}}}", "")
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    sp = shape._element
                    sp.getparent().remove(sp)
                    slide = shape.part.slide
                    slide.shapes.add_picture(img_path, left, top, width=width, height=height)
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

def process_shape(shape, row, images_dir):
    """Process shapes and tables recursively—image first."""
    try:
        replace_images_on_shape(shape, row, images_dir)  # Image first
        replace_text_in_obj(shape, row)  # Text second
        if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
            for row_cells in shape.table.rows:
                for cell in row_cells.cells:
                    replace_images_on_shape(cell, row, images_dir)
                    replace_text_in_obj(cell, row)
    except Exception as e:
        logger.error(f"Error in process_shape: {str(e)}")

@app.post("/api/generate")
async def generate(excel: UploadFile = File(...), ppt: UploadFile = File(...), images: UploadFile = File(None)):
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            # Save uploaded files
            excel_filename = excel.filename or "content.xlsx"
            ppt_filename = ppt.filename or "template_client.pptx"
            excel_path = os.path.join(tmpdir, excel_filename)
            ppt_path = os.path.join(tmpdir, ppt_filename)
            images_dir = os.path.join(tmpdir, "images")
            logger.info(f"Saving files: excel={excel_path}, ppt={ppt_path}")

            # Validate and save files
            excel_content = await excel.read()
            if not excel_content:
                logger.error("Excel file is empty")
                raise HTTPException(status_code=400, detail="Excel file is empty")
            with open(excel_path, "wb") as f:
                f.write(excel_content)

            ppt_content = await ppt.read()
            if not ppt_content:
                logger.error("PowerPoint file is empty")
                raise HTTPException(status_code=400, detail="PowerPoint file is empty")
            with open(ppt_path, "wb") as f:
                f.write(ppt_content)

            if images:
                zip_filename = images.filename or "images.zip"
                zip_path = os.path.join(tmpdir, zip_filename)
                zip_content = await images.read()
                if not zip_content:
                    logger.error("Images ZIP file is empty")
                    raise HTTPException(status_code=400, detail="Images ZIP file is empty")
                with open(zip_path, "wb") as f:
                    f.write(zip_content)
                try:
                    with zipfile.ZipFile(zip_path
