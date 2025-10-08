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
import aiofiles  # Async file I/O for speed

# Setup logging (reduced in production)
logging.basicConfig(level=logging.WARNING)  # Less verbose
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
    """Return a safe string value for a field from Excel."""
    try:
        if field in row and not pd.isna(row[field]):
            return str(row[field])
        return ""
    except Exception as e:
        logger.error(f"Error in get_value_for_field for field {field}: {str(e)}")
        return ""

def find_image_path(value, images_dir):
    """Try to resolve image path in images_dir."""
    try:
        if not value or not images_dir:
            return None
        candidate = os.path.join(images_dir, value)
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
    """Process shapes and tables recursively."""
    try:
        replace_text_in_obj(shape, row)
        replace_images_on_shape(shape, row, images_dir)
        if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
            for row_cells in shape.table.rows:
                for cell in row_cells.cells:
                    replace_text_in_obj(cell, row)
    except Exception as e:
        logger.error(f"Error in process_shape: {str(e)}")

@app.post("/api/generate")
async def generate(excel: UploadFile = File(...), ppt: UploadFile = File(...), images: UploadFile = File(None)):
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            # Save uploaded files (async for speed)
            excel_filename = excel.filename or "content.xlsx"
            ppt_filename = ppt.filename or "template_client.pptx"
            excel_path = os.path.join(tmpdir, excel_filename)
            ppt_path = os.path.join(tmpdir, ppt_filename)
            images_dir = os.path.join(tmpdir, "images")

            # Async save Excel
            excel_content = await excel.read()
            if not excel_content:
                raise HTTPException(status_code=400, detail="Excel file is empty")
            async with aiofiles.open(excel_path, "wb") as f:
                await f.write(excel_content)

            # Async save PPT
            ppt_content = await ppt.read()
            if not ppt_content:
                raise HTTPException(status_code=400, detail="PowerPoint file is empty")
            async with aiofiles.open(ppt_path, "wb") as f:
                await f.write(ppt_content)

            # Async save ZIP if present
            if images:
                zip_filename = images.filename or "images.zip"
                zip_path = os.path.join(tmpdir, zip_filename)
                zip_content = await images.read()
                if not zip_content:
                    raise HTTPException(status_code=400, detail="Images ZIP file is empty")
                async with aiofiles.open(zip_path, "wb") as f:
                    await f.write(zip_content)
                try:
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(images_dir)
                except zipfile.BadZipFile:
                    images_dir = None

            # Load Excel and PowerPoint
            df = pd.read_excel(excel_path)
            prs = Presentation(ppt_path)

            # Process slides (original)
            if len(df) > len(prs.slides):
                logger.warning("More rows in Excel than slides in template. Extra rows ignored.")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    process_shape(shape, row, images_dir if images_dir else tmpdir)

            # Save output
            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            prs.save(output_file)

            # Verify output file
            if not os.path.exists(output_file) or os.path.getsize(output_file) == 0:
                raise HTTPException(status_code=500, detail="Output presentation file is missing or empty")

            # Async read for streaming
            async with aiofiles.open(output_file, "rb") as f:
                file_content = await f.read()

            return StreamingResponse(
                io.BytesIO(file_content),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": "attachment; filename=Client_Presentation.pptx"}
            )
        except Exception as e:
            logger.error(f"Error in /api/generate: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")
