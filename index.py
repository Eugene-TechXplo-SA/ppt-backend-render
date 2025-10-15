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
    """Return a safe string value for a field from Excel."""
    try:
        if field in row and not pd.isna(row[field]):
            return str(row[field])
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
        # Check if it's an image file
        if os.path.exists(candidate) and os.path.splitext(candidate)[1].lower() in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
            return candidate
        logger.warning(f"Image not found or invalid: {candidate}")
        return None
    except Exception as e:
        logger.error(f"Error in find_image_path for {value}: {str(e)}")
        return None

def replace_images_on_shape(shape, row, images_dir):
    """Replace placeholders with images if the value points to a file."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if not (hasattr(shape, "has_text_frame") and shape.has_text_frame):
            return
        if hasattr(shape, "has_table") and shape.has_table:
            return  # Skip tables; images in cells not supported natively
        full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
        matches = placeholder_pattern.findall(full_text)
        for field in matches:
            img_path = find_image_path(get_value_for_field(row, field), images_dir)
            if img_path:
                # Clear all placeholders in the shape first
                clear_pattern = re.compile(r"\{\{(.*?)\}\}")
                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        r.text = clear_pattern.sub("", r.text)
                # Then remove and replace the shape with image
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                sp = shape._element
                sp.getparent().remove(sp)
                slide = shape.part.slide
                slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                logger.info(f"Replaced shape with image: {img_path}")
                break  # Only handle one image per shape
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

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
                        if field.strip().lower() == "link" and val:
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

def process_shape(shape, row, images_dir):
    """Process shapes and tables recursively."""
    try:
        # First: Handle images (before text replacement removes placeholders)
        replace_images_on_shape(shape, row, images_dir)
        # Then: Handle remaining text
        replace_text_in_obj(shape, row)
        if hasattr(shape, "shape_type") and shape.shape_type == 19:  # TABLE
            for table_row in shape.table.rows:
                for cell in table_row.cells:
                    # For cells: images not supported, so only text
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

            images_dir = None  # Default to None
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
                        # Security: Prevent path traversal
                        for name in zip_ref.namelist():
                            if '..' in name or name.startswith('/'):
                                raise HTTPException(status_code=400, detail="Invalid ZIP contents: Path traversal detected")
                        zip_ref.extractall(images_dir)
                    logger.info(f"Extracted images to: {images_dir}")
                    images_dir = images_dir  # Set if successful
                except zipfile.BadZipFile as e:
                    logger.warning(f"Invalid ZIP file: {str(e)}")
                    images_dir = None

            # Load Excel and PowerPoint
            logger.info("Loading Excel and PowerPoint")
            try:
                df = pd.read_excel(excel_path, sheet_name=0)
                if df.empty:
                    raise ValueError("No data found in Excel sheet")
            except Exception as e:
                logger.error(f"Failed to load Excel: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid Excel file: {str(e)}")
            try:
                prs = Presentation(ppt_path)
            except Exception as e:
                logger.error(f"Failed to load PowerPoint: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid PowerPoint file: {str(e)}")

            # Process slides
            logger.info("Processing slides")
            num_rows = len(df)
            num_slides = len(prs.slides)
            if num_rows > num_slides:
                logger.warning(f"More rows in Excel ({num_rows}) than slides in template ({num_slides}). Extra rows will be ignored.")
            elif num_rows < num_slides:
                logger.warning(f"Fewer rows in Excel ({num_rows}) than slides ({num_slides}). Later slides will use no data.")
            for i, row in df.iterrows():
                if i >= num_slides:
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    process_shape(shape, row, images_dir)

            # Save output
            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            try:
                prs.save(output_file)
                logger.info(f"Saved output to: {output_file}")
            except Exception as e:
                logger.error(f"Failed to save PowerPoint: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Failed to save presentation: {str(e)}")

            # Verify output file
            if not os.path.exists(output_file):
                logger.error("Output file does not exist")
                raise HTTPException(status_code=500, detail="Output presentation file is missing")
            if os.path.getsize(output_file) == 0:
                logger.error("Output file is empty")
                raise HTTPException(status_code=500, detail="Output presentation file is empty")

            # Stream directly without loading into memory
            logger.info("Returning StreamingResponse")
            def file_stream():
                with open(output_file, "rb") as f:
                    while chunk := f.read(8192):
                        yield chunk

            return StreamingResponse(
                file_stream(),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": "attachment; filename=Client_Presentation.pptx"}
            )
        except Exception as e:
            logger.error(f"Error in /api/generate: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")
