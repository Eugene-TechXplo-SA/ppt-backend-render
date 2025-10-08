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
from copy import deepcopy

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
    """Replace text placeholders inside shape or cell—scans full text."""
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            full_text = "".join(run.text for para in obj.text_frame.paragraphs for run in para.runs)
            matches = placeholder_pattern.findall(full_text)
            for field in matches:
                val = get_value_for_field(row, field)
                # Replace in all runs
                for para in obj.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = run.text.replace(f"{{{{{field}}}}}", val)
                        if field.lower() == "link" and val:
                            run.font.color.rgb = RGBColor(0, 0, 255)
                            run.font.underline = True
            logger.info(f"Replaced {len(matches)} placeholders in shape")
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
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            for table_row in shape.table.rows:
                for cell in table_row.cells:
                    replace_text_in_obj(cell, row)
    except Exception as e:
        logger.error(f"Error in process_shape: {str(e)}")

def duplicate_slide_exact(prs, template_slide_index):
    """Exact duplicate of slide using deepcopy on part."""
    try:
        template_slide = prs.slides[template_slide_index]
        # Deep copy the slide part
        new_slide_part = deepcopy(template_slide.part)
        # Add the new slide to prs
        new_slide = prs.slides.add_slide(template_slide.slide_layout)
        new_slide.part = new_slide_part
        # Copy relationships (for images/text)
        new_slide.part.rels = deepcopy(template_slide.part.rels)
        logger.info(f"Exact duplicated slide {template_slide_index}")
        return new_slide
    except Exception as e:
        logger.error(f"Failed to exact duplicate slide: {str(e)}")
        return None

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
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(images_dir)
                    logger.info(f"Extracted images to: {images_dir}")
                except zipfile.BadZipFile as e:
                    logger.warning(f"Invalid ZIP file: {str(e)}")
                    images_dir = None

            # Load Excel and PowerPoint
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

            # Step 1: Duplicate slides BEFORE processing (exact clone with {{}})
            num_rows = len(df)
            template_slide_index = 0
            if num_rows > 1:
                for i in range(1, num_rows):
                    try:
                        duplicate_slide_exact(prs, template_slide_index)
                        logger.info(f"Exact duplicated slide for row {i}")
                    except Exception as e:
                        logger.warning(f"Skipped duplicating slide {i}: {str(e)}")

            # Step 2: Fill data in ALL slides (now each has {{}} from clone)
            logger.info("Filling data in slides")
            num_slides = len(prs.slides)
            for i, row in df.iterrows():
                if i < num_slides:
                    slide = prs.slides[i]
                    for shape in slide.shapes:
                        process_shape(shape, row, images_dir if images_dir else tmpdir)
                else:
                    logger.warning(f"Extra row {i} ignored—no slide")
                    break

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

            # Read file for streaming
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
