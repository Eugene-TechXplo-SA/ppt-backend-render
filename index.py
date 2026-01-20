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
from typing import Dict, List, Optional, Tuple
import unicodedata

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

def normalize_name(name: str) -> str:
    """Normalize a name for flexible matching: lowercase, remove special chars, normalize spaces."""
    if not name:
        return ""
    normalized = unicodedata.normalize('NFKD', name)
    normalized = normalized.encode('ASCII', 'ignore').decode('ASCII')
    normalized = re.sub(r'[^a-zA-Z0-9\s]', '', normalized)
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip().lower()

def find_folder_for_site(site_identifier: str, images_dir: str) -> Optional[str]:
    """Find a folder matching the site identifier using flexible matching."""
    if not site_identifier or not os.path.exists(images_dir):
        return None

    normalized_site = normalize_name(site_identifier)

    for item in os.listdir(images_dir):
        item_path = os.path.join(images_dir, item)
        if os.path.isdir(item_path):
            normalized_folder = normalize_name(item)
            if normalized_folder == normalized_site:
                logger.info(f"Matched site '{site_identifier}' to folder '{item}'")
                return item_path

    return None

def get_first_image_from_folder(folder_path: str) -> Optional[str]:
    """Get the first image file from a folder (alphabetically)."""
    if not os.path.exists(folder_path):
        return None

    image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp'}
    images = []

    for file in os.listdir(folder_path):
        file_lower = file.lower()
        if any(file_lower.endswith(ext) for ext in image_extensions):
            images.append(file)

    if images:
        images.sort()
        return os.path.join(folder_path, images[0])

    return None

def find_image_path_enhanced(value: str, images_dir: str, site_identifier: Optional[str] = None) -> Optional[str]:
    """
    Enhanced image finding with folder-based support.

    Strategy:
    1. If site_identifier provided, look for folder matching site name
    2. Fall back to exact filename matching (backward compatibility)
    3. Fall back to case-insensitive filename search
    """
    try:
        if site_identifier:
            folder_path = find_folder_for_site(site_identifier, images_dir)
            if folder_path:
                image_path = get_first_image_from_folder(folder_path)
                if image_path:
                    logger.info(f"Found image in folder: {image_path}")
                    return image_path

        clean_value = value.replace("images/", "", 1) if value.startswith("images/") else value
        candidate = os.path.join(images_dir, clean_value)
        if os.path.exists(candidate):
            logger.info(f"Found image by exact path: {candidate}")
            return candidate

        for root, dirs, files in os.walk(images_dir):
            for file in files:
                if file.lower() == clean_value.lower():
                    found_path = os.path.join(root, file)
                    logger.info(f"Found image by case-insensitive search: {found_path}")
                    return found_path

        logger.warning(f"No image found for value: {value}")
        return None
    except Exception as e:
        logger.error(f"Error in find_image_path_enhanced: {str(e)}")
        return None

def validate_zip_structure(zip_path: str, df: pd.DataFrame, site_column: Optional[str] = None) -> Dict:
    """Validate ZIP structure and return detailed report."""
    validation_result = {
        "valid": True,
        "errors": [],
        "warnings": [],
        "structure": {
            "has_folders": False,
            "folders": [],
            "loose_files": [],
            "matched_sites": [],
            "unmatched_sites": [],
            "extra_folders": []
        }
    }

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            extract_dir = os.path.join(tmpdir, "extracted")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)

            items = os.listdir(extract_dir)
            folders = [item for item in items if os.path.isdir(os.path.join(extract_dir, item))]
            files = [item for item in items if os.path.isfile(os.path.join(extract_dir, item))]

            validation_result["structure"]["has_folders"] = len(folders) > 0
            validation_result["structure"]["folders"] = folders
            validation_result["structure"]["loose_files"] = files

            if site_column and site_column in df.columns:
                sites = df[site_column].dropna().unique().tolist()
                sites = [str(s).strip() for s in sites]

                for site in sites:
                    folder_path = find_folder_for_site(site, extract_dir)
                    if folder_path:
                        image = get_first_image_from_folder(folder_path)
                        if image:
                            validation_result["structure"]["matched_sites"].append(site)
                        else:
                            validation_result["structure"]["unmatched_sites"].append(site)
                            validation_result["warnings"].append(f"Folder found for '{site}' but no images inside")
                    else:
                        validation_result["structure"]["unmatched_sites"].append(site)
                        validation_result["warnings"].append(f"No folder found for site: {site}")

                normalized_sites = {normalize_name(s) for s in sites}
                for folder in folders:
                    if normalize_name(folder) not in normalized_sites:
                        validation_result["structure"]["extra_folders"].append(folder)
                        validation_result["warnings"].append(f"Extra folder not referenced in Excel: {folder}")

            if len(folders) == 0 and len(files) == 0:
                validation_result["valid"] = False
                validation_result["errors"].append("ZIP file is empty")

    except Exception as e:
        validation_result["valid"] = False
        validation_result["errors"].append(f"Error validating ZIP: {str(e)}")

    return validation_result

def get_value_for_field(row, field):
    try:
        if field in row and not pd.isna(row[field]):
            return str(row[field]).strip()
        return ""
    except:
        return ""

def is_image_path(val):
    return val and isinstance(val, str) and bool(re.search(r'\.(jpg|jpeg|png|gif|bmp)$', val.lower()))

def replace_text_in_obj(obj, row, images_dir, site_identifier=None):
    placeholder_pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")
    try:
        if hasattr(obj, "text_frame") and obj.text_frame is not None:
            for paragraph in obj.text_frame.paragraphs:
                for run in paragraph.runs:
                    matches = placeholder_pattern.findall(run.text)
                    for field_raw in matches:
                        field_clean = re.sub(r'\s+', ' ', field_raw.strip())
                        field_key = field_clean.lower()

                        col = next((c for c in row.index if c.strip().lower() == field_key), None)
                        if not col:
                            continue
                        val = get_value_for_field(row, col)

                        if is_image_path(val):
                            continue

                        escaped = re.escape(f"{{{{{field_raw}}}}}")
                        flexible = (
                            escaped
                            .replace('\\{\\{', r'\s*\{\{\s*')
                            .replace('\\}\\}', r'\s*\}\}')
                            .replace('\\-', r'[-–—]')
                            .replace('\\ ', r'\s*')
                        )
                        match = re.search(flexible, run.text, re.UNICODE)
                        if not match:
                            continue
                        placeholder_text = match.group(0)

                        if col.lower().endswith("link"):
                            try:
                                result = urlparse(val)
                                if all([result.scheme, result.netloc]):
                                    run.text = run.text.replace(placeholder_text, val, 1)
                                    run.font.color.rgb = RGBColor(0, 0, 255)
                                    run.font.underline = True
                                    continue
                            except:
                                pass

                        run.text = run.text.replace(placeholder_text, val, 1)

    except Exception as e:
        logger.error(f"Error in replace_text_in_obj: {str(e)}")

def replace_images_on_shape(shape, row, images_dir, site_identifier=None):
    placeholder_pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")
    try:
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            full_text = "".join(r.text or "" for p in shape.text_frame.paragraphs for r in p.runs)
            matches = placeholder_pattern.findall(full_text)
            for field_raw in matches:
                field_clean = re.sub(r'\s+', ' ', field_raw.strip())
                field_key = field_clean.lower()
                col = next((c for c in row.index if c.strip().lower() == field_key), None)
                if not col:
                    continue
                val = get_value_for_field(row, col)
                if is_image_path(val):
                    img_path = find_image_path_enhanced(val, images_dir, site_identifier)
                    if img_path:
                        for p in shape.text_frame.paragraphs:
                            for r in p.runs:
                                r.text = r.text.replace(f"{{{{{field_raw}}}}}", "", 1)
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        sp = shape._element
                        sp.getparent().remove(sp)
                        slide = shape.part.slide
                        slide.shapes.add_picture(img_path, left, top, width=width, height=height)
                        return
    except Exception as e:
        logger.error(f"Error in replace_images_on_shape: {str(e)}")

@app.post("/api/validate")
async def validate(excel: UploadFile = File(...), images: UploadFile = File(None)):
    """Validate uploaded files before processing."""
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            excel_path = os.path.join(tmpdir, excel.filename or "content.xlsx")
            with open(excel_path, "wb") as f:
                f.write(await excel.read())

            df = pd.read_excel(excel_path)
            df.columns = [col.strip() for col in df.columns]

            result = {
                "valid": True,
                "excel": {
                    "rows": len(df),
                    "columns": list(df.columns),
                    "sample_data": df.head(3).to_dict(orient="records")
                },
                "images": None
            }

            if images:
                zip_path = os.path.join(tmpdir, images.filename or "images.zip")
                with open(zip_path, "wb") as f:
                    f.write(await images.read())

                site_column = None
                for col in df.columns:
                    if 'site' in col.lower() or 'name' in col.lower() or 'id' in col.lower():
                        site_column = col
                        break

                validation = validate_zip_structure(zip_path, df, site_column)
                result["images"] = validation
                result["valid"] = validation["valid"]

            return result

        except Exception as e:
            logger.error(f"Error in /api/validate: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Validation Error: {str(e)}")

@app.post("/api/generate")
async def generate(
    excel: UploadFile = File(...),
    ppt: UploadFile = File(...),
    images: UploadFile = File(None),
    site_column: Optional[str] = None
):
    """Generate PowerPoint with enhanced folder-based image support."""
    with tempfile.TemporaryDirectory() as tmpdir:
        try:
            excel_path = os.path.join(tmpdir, excel.filename or "content.xlsx")
            ppt_path = os.path.join(tmpdir, ppt.filename or "template.pptx")
            images_dir = os.path.join(tmpdir, "images")

            with open(excel_path, "wb") as f:
                f.write(await excel.read())
            with open(ppt_path, "wb") as f:
                f.write(await ppt.read())

            if images:
                zip_path = os.path.join(tmpdir, images.filename or "images.zip")
                with open(zip_path, "wb") as f:
                    f.write(await images.read())
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(images_dir)
                logger.info(f"Extracted images to: {images_dir}")

            df = pd.read_excel(excel_path)
            df.columns = [col.strip() for col in df.columns]

            if not site_column:
                for col in df.columns:
                    if 'site' in col.lower() or 'name' in col.lower():
                        site_column = col
                        break

            prs = Presentation(ppt_path)

            logger.info("Processing slides – Step 1: Images")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]

                site_id = None
                if site_column and site_column in row.index:
                    site_id = get_value_for_field(row, site_column)

                for shape in slide.shapes:
                    replace_images_on_shape(shape, row, images_dir if images else tmpdir, site_id)
                    if hasattr(shape, "table"):
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_images_on_shape(cell, row, images_dir if images else tmpdir, site_id)

            logger.info("Processing slides – Step 2: Text")
            for i, row in df.iterrows():
                if i >= len(prs.slides):
                    break
                slide = prs.slides[i]

                site_id = None
                if site_column and site_column in row.index:
                    site_id = get_value_for_field(row, site_column)

                for shape in slide.shapes:
                    replace_text_in_obj(shape, row, images_dir if images else tmpdir, site_id)
                    if hasattr(shape, "table"):
                        for row_cells in shape.table.rows:
                            for cell in row_cells.cells:
                                replace_text_in_obj(cell, row, images_dir if images else tmpdir, site_id)

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

@app.get("/health")
async def health():
    return {"status": "ok"}
