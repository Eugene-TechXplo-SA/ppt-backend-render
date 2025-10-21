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
            val = str(row[field]).strip()
            return val
        return ""
    except Exception as e:
        logger.error(f"Error in get_value_for_field for {field}: {str(e)}")
        return ""

def is_image_value(val):
    """Check if value looks like an image path—expanded extensions."""
    if val and isinstance(val, str):
        return bool(re.search(r'\.(jpg|jpeg|png|gif|bmp|svg|tiff|webp)$', val.lower()))
    return False

def find_image_path(value, images_dir):
    """Try to resolve image path in images_dir—handle casing variants."""
    try:
        if not value or not images_dir:
            logger.warning(f"No value or images_dir for {value}")
            return None
        clean_value = re.sub(r'(?i)^images/', '', value.strip())
        base_dir = os.path.dirname(images_dir) if os.path.isdir(images_dir) else images_dir
        logger.debug(f"Searching for {clean_value} in {base_dir}")
        for dir_var in ["images", "Images", "IMAGEs"]:
            candidate = os.path.join(base_dir, dir_var, clean_value)
            if os.path.exists(candidate):
                logger.info(f"Found image: {candidate}")
                return candidate
        logger.warning(f"Image not found for {value} in {images_dir}")
        return None
    except Exception as e:
        logger.error(f"Error in find_image_path for {value}: {str(e)}")
        return None

def replace_text_in_obj(obj, row):
    """Replace text placeholders inside shape or cell—case-insensitive match."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field in matches:
                        field_lower = field.lower().strip()
                        matching_col = next((col for col in row.index if col.lower().strip() == field_lower), None)
                        if matching_col:
                            val = get_value_for_field(row, matching_col)
                            run.text = run.text.replace(f"{{{{{field}}}}}", val)
                            if field_lower == "link" and val:
                                run.font.color.rgb = RGBColor(0, 0, 255)
                                run.font.underline = True
                        else:
                            logger.warning(f"No matching column for {field}")
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

def replace_images_on_shape(shape, row, images_dir):
    """Replace placeholders with images—case-insensitive, no 'image' filter."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            replaced = True
            while replaced:
                replaced = False
                full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
                matches = placeholder_pattern.findall(full_text)
                for field in matches:
                    field_lower = field.lower().strip()
                    matching_col = next((col for col in row.index if col.lower().strip() == field_lower), None)
                    if matching_col:
                        img_value = get_value_for_field(row, matching_col)
                        if is_image_value(img_value):
                            img_path = find_image_path(img_value, images_dir)
                            if img_path:
                                logger.info(f"Inserting image for {field} from {img_path}")
                                for p in shape.text_frame.paragraphs:
                                    for r in p.runs:
                                        r.text = r.text.replace(f"{{{{{field}}}}}", "")  # Clear placeholder
                                try:
                                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                                    sp = shape._element
                                    sp.getparent().remove(sp)
                                    slide = shape.part.slide
                                    slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                                    logger.info(f"Image inserted: {img_path}")
                                    replaced = True
                                    break
                                except Exception as e:
                                    logger.error(f"Failed to insert image at {img_path}: {str(e)}")
                                    # Fallback: Clear text if insertion fails
                                    for p in shape.text_frame.paragraphs:
                                        p.text = ""
                            else:
                                logger.warning(f"No image path found for {field}: {img_value}")
                    else:
                        logger.warning(f"No matching column for {field}")
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

def replace_text_in_obj(obj, row):
    """Replace remaining text placeholders inside shape or cell—case-insensitive match."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field in matches:
                        field_lower = field.lower().strip()
                        matching_col = next((col for col in row.index if col.lower().strip() == field_lower), None)
                        if matching_col:
                            val = get_value_for_field(row, matching_col)
                            run.text = run.text.replace(f"{{{{{field}}}}}", val)
                            if field_lower == "link" and val:
                                run.font.color.rgb = RGBColor(0, 0, 255)
                                run.font.underline = True
                        else:
                            logger.warning(f"No match for {field}")
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

def process_shape(shape, row, images_dir):
    """Process shapes—image first, then text."""
    try:
        replace_images_on_shape(shape, row, images_dir)  # Images first—no text left
        replace_text_in_obj(shape, row)  # Text second—for non-image fields
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
