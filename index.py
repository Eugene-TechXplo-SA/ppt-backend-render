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

# Setup logging to track errors and info (useful for debugging)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# CORS middleware to allow frontend calls (e.g., from Bolt.new or Netlify)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For testing; restrict in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Function to get value from Excel row (trim spaces for clean data)
def get_value_for_field(row, field):
    try:
        if field in row and not pd.isna(row[field]):
            val = str(row[field]).strip()
            return val
        return ""
    except Exception as e:
        logger.error(f"Error getting value for {field}: {str(e)}")
        return ""

# Function to check if a value looks like an image path (ends with .jpg etc.)
def is_image_value(val):
    if val and isinstance(val, str):
        return bool(re.search(r'\.(jpg|jpeg|png|gif|bmp)$', val.lower()))
    return False

# Function to find image file in extracted ZIP (strip 'images/' prefix)
def find_image_path(value, images_dir):
    try:
        if not value or not images_dir:
            return None
        clean_value = value.replace("images/", "", 1) if value.startswith("images/") else value
        candidate = os.path.join(images_dir, clean_value)
        if os.path.exists(candidate):
            logger.info(f"Found image: {candidate}")
            return candidate
        logger.warning(f"Image not found: {candidate}")
        return None
    except Exception as e:
        logger.error(f"Error finding image for {value}: {str(e)}")
        return None

# Function to replace text placeholders with Excel data
def replace_text_in_obj(obj, row):
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
        logger.error(f"Error replacing text: {str(e)}")

# Function to replace image placeholders with actual images from ZIP
def replace_images_on_shape(shape, row, images_dir):
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            replaced = True
            while replaced:
                replaced = False
                full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
                matches = placeholder_pattern.findall(full_text)
                for field in matches:
                    val = get_value_for_field(row, field)
                    if is_image_value(val):
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
                            logger.info(f"Inserted image: {img_path}")
                            replaced = True
                            break
    except Exception as e:
        logger.error(f"Error replacing images: {str(e)}")

# Function to process each shape (image first, then text)
def process_shape(shape, row, images_dir):
    try:
        replace_images_on_shape(shape, row, images_dir)  # Run images first (complete before text)
        replace_text_in_obj(shape, row)  # Run text second (only after images are done)
        if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
            for row_cells in shape.table.rows:
                for cell in row_cells.cells:
                    replace_images_on_shape(cell, row, images_dir)  # Images in tables
                    replace_text_in_obj(cell, row)  # Text in tables
    except Exception as e:
        logger.error(f"Error processing shape: {str(e)}")

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
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(images_dir)
                    logger.info(f"Extracted images to: {images_dir}")
                except zipfile.BadZipFile as e:
                    logger.warning(f"Invalid ZIP file: {str(e)}")
                    images_dir = None

            logger.info("Loading Excel and PowerPoint")
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                logger.error(f"Failed to load Excel: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid Excel file: {str(e)}")
            try:
                prs = Presentation(ppt_path)
            except Exception as e:
                logger.error(f"Failed to load PowerPoint: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid PowerPoint file: {str(e)}")

            logger.info("Processing slides")
            if len(df) > len(prs.slides):
                logger.warning("More rows in Excel than slides in template. Extra rows will be ignored.")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    process_shape(shape, row, images_dir if images_dir else tmpdir)

            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            try:
                prs.save(output_file)
                logger.info(f"Saved output to: {output_file}")
            except Exception as e:
                logger.error(f"Failed to save PowerPoint: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Failed to save presentation: {str(e)}")

            if not os.path.exists(output_file):
                logger.error("Output file does not exist")
                raise HTTPException(status_code=500, detail="Output presentation file is missing")
            if os.path.getsize(output_file) == 0:
                logger.error("Output file is empty")
                raise HTTPException(status_code=500, detail="Output presentation file is empty")

            logger.info("Reading output file for streaming")
            try:
                with open(output_file, "rb") as f:
                    file_content = f.read()
                if not file_content:
                    logger.error("Output file is empty when read")
                    raise HTTPException(status_code=500, detail="Output file is empty when read")
            except Exception as e:
                logger.error(f"Failed to read output file: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Failed to read output file: {str(e)}")

            logger.info("Returning StreamingResponse")
            return StreamingResponse(
                io.BytesIO(file_content),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": "attachment; filename=Client_Presentation.pptx"}
            )
        except Exception as e:
            logger.error(f"Error in /api/generate: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")
