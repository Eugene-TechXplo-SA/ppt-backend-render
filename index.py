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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_value_for_field(row, field):
    try:
        if field in row and not pd.isna(row[field]):
            return str(row[field]).strip()
        return ""
    except Exception as e:
        logger.error(f"Error in get_value_for_field for {field}: {str(e)}")
        return ""

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
        logger.error(f"Error in find_image_path for {value}: {str(e)}")
        return None

def replace_text_in_obj(obj, row):
    placeholder_pattern = re.compile(r"\{\{(.*?)\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for p in obj.text_frame.paragraphs:
                for r in p.runs:
                    matches = placeholder_pattern.findall(r.text)
                    for field in matches:
                        field_lower = field.lower().strip()
                        matching_col = next((col for col in row.index if col.lower().strip() == field_lower), None)
                        if matching_col:
                            val = get_value_for_field(row, matching_col)
                            r.text = r.text.replace(f"{{{{{field}}}}}", val)
                            if field_lower == "link" and val:
                                r.font.color.rgb = RGBColor(0, 0, 255)
                                r.font.underline = True
                        else:
                            logger.warning(f"No match for {field}")
    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

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
                    field_lower = field.lower().strip()
                    matching_col = next((col for col in row.index if col.lower().strip() == field_lower), None)
                    if matching_col:
                        img_value = get_value_for_field(row, matching_col)
                        img_path = find_image_path(img_value, images_dir)
                        if img_path:
                            for p in shape.text_frame.paragraphs:
                                for r in p.runs:
                                    r.text = r.text.replace(f"{{{{{field}}}}}", "")
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            sp = shape._element
                            sp.getparent().remove(sp)
                            slide = shape.part.slide
                            slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                            logger.info(f"Inserted: {img_path}")
                            replaced = True
                            break
                    else:
                        logger.warning(f"No match for image {field}")
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

def process_shape(shape, row, images_dir):
    try:
        replace_images_on_shape(shape, row, images_dir)
        replace_text_in_obj(shape, row)
        if hasattr(shape, "shape_type") and shape.shape_type == 19:
            for r in shape.table.rows:
                for c in r.cells:
                    replace_images_on_shape(c, row, images_dir)
                    replace_text_in_obj(c, row)
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
            logger.info(f"Saving: excel={excel_path}, ppt={ppt_path}")

            excel_content = await excel.read()
            if not excel_content:
                raise HTTPException(status_code=400, detail="Excel empty")
            with open(excel_path, "wb") as f:
                f.write(excel_content)

            ppt_content = await ppt.read()
            if not ppt_content:
                raise HTTPException(status_code=400, detail="PPT empty")
            with open(ppt_path, "wb") as f:
                f.write(ppt_content)

            if images:
                zip_filename = images.filename or "images.zip"
                zip_path = os.path.join(tmpdir, zip_filename)
                zip_content = await images.read()
                if not zip_content:
                    raise HTTPException(status_code=400, detail="ZIP empty")
                with open(zip_path, "wb") as f:
                    f.write(zip_content)
                try:
                    with zipfile.ZipFile(zip_path, "r") as zip_ref:
                        zip_ref.extractall(images_dir)
                    logger.info(f"Extracted to: {images_dir}")
                except zipfile.BadZipFile as e:
                    logger.warning(f"Invalid ZIP: {str(e)}")
                    images_dir = None

            logger.info("Loading files")
            try:
                df = pd.read_excel(excel_path)
            except Exception as e:
                logger.error(f"Excel load failed: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid Excel: {str(e)}")
            try:
                prs = Presentation(ppt_path)
            except Exception as e:
                logger.error(f"PPT load failed: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid PPT: {str(e)}")

            logger.info("Processing")
            if len(df) > len(prs.slides):
                logger.warning("Extra Excel rows ignored")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]
                for shape in slide.shapes:
                    process_shape(shape, row, images_dir if images_dir else tmpdir)

            output_file = os.path.join(tmpdir, "Client_Presentation.pptx")
            try:
                prs.save(output_file)
                logger.info(f"Saved to: {output_file}")
            except Exception as e:
                logger.error(f"Save failed: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Save failed: {str(e)}")

            if not os.path.exists(output_file):
                logger.error("Output missing")
                raise HTTPException(status_code=500, detail="Output missing")
            if os.path.getsize(output_file) == 0:
                logger.error("Output empty")
                raise HTTPException(status_code=500, detail="Output empty")

            logger.info("Reading output")
            try:
                with open(output_file, "rb") as f:
                    file_content = f.read()
                if not file_content:
                    logger.error("Output read empty")
                    raise HTTPException(status_code=500, detail="Output read empty")
            except Exception as e:
                logger.error(f"Read failed: {str(e)}")
                raise HTTPException(status_code=500, detail=f"Read failed: {str(e)}")

            logger.info("Returning response")
            return StreamingResponse(
                io.BytesIO(file_content),
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                headers={"Content-Disposition": "attachment; filename=Client_Presentation.pptx"}
            )
        except Exception as e:
            logger.error(f"Generate failed: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Internal error: {str(e)}")
