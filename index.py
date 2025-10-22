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
        # Try case-insensitive search as a fallback
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

def replace_images_on_shape(shape, row, images_dir):
    """Replace placeholders with images if the value is a path."""
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
                        logger.info(f"Attempting to insert image for {field} from {img_path}")
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
                    else:
                        logger.error(f"Failed to find image path for {field}: {val}")
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

def replace_text_in_obj(obj, row):
    """Replace placeholders with text, handling links specially."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field in matches:
                        val = get_value_for_field(row, field)
                        if not is_image_path(val):  # Skip if it’s an image path
                            if val and field.lower().endswith("link"):  # Handle links
                                try:
                                    result = urlparse(val)
                                    if all([result.scheme, result.netloc]):  # Valid URL
                                        run.text = run.text.replace(f"{{{{{field}}}}}", val)
                                        run.font.color.rgb = RGBColor(0, 0, 255)  # Blue
                                        run.font.underline = True  # Underline
                                        logger.info(f"Replaced {field} with hyperlink: {val}")
                                    else:
                                        run.text = run.text.replace(f"{{{{{field}}}}}", val)
                                        logger.warning(f"Invalid URL for {field}: {val}")
                                except ValueError:
                                    run.text = run.text.replace(f"{{{{{field}}}}}", val)
                                    logger.warning(f"Invalid URL format for {field}: {val}")
                            else:
                                run.text = run.text.replace(f"{{{{{field}}}}}", val)
                                logger.info(f"Replaced {field} with text: {val}")
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

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
                raise HTTPException(status_code=400, detail="Excel file is empty")
            with open(excel_path, "wb") as f:
                f.write(excel_content)

            ppt_content = await ppt.read()
            if not ppt_content:
                raise HTTPException(status_code=400, detail="PowerPoint file is empty")
            with open(ppt_path, "wb") as f:
                f.write(ppt_content)

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
                    extracted_files = [f.filename for f in zip_ref.infolist()]
                    logger.info(f"Extracted files: {extracted_files}")
                except zipfile.BadZipFile as e:
                    logger.error(f"Invalid ZIP file: {str(e)}")
                    raise HTTPException(status_code=400, detail=f"Invalid ZIP file: {str(e)}")

            logger.info("Loading Excel and PowerPoint")
            try:
                df = pd.read_excel(excel_path)
                # Trim column names to handle trailing spaces
                df.columns = [col.strip() for col in df.columns]
            except Exception as e:
                logger.error(f"Failed to load Excel: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid Excel file: {str(e)}")
            try:
                prs = Presentation(ppt_path)
            except Exception as e:
                logger.error(f"Failed to load PowerPoint: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid PowerPoint file: {str(e)}")

            logger.info("Processing slides - Step 1: Images")
            # First pass: Process all images across all slides
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_images_on_shape(shape, row, images_dir if images_dir else tmpdir)
                    if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_images_on_shape(cell, row, images_dir if images_dir else tmpdir)

            logger.info("Processing slides - Step 2: Text")
            # Second pass: Process all text across all slides
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    replace_text_in_obj(shape, row)
                    if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_text_in_obj(cell, row)

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
