from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
import zipfile
import pandas as pd
from pptx import Presentation
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
                        left, top, width, height = shape.left
