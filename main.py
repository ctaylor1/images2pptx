import logging
import sys
import os
from pathlib import Path
import yaml
import loguru
from loguru import logger
from pptx import Presentation
from pptx.util import Inches, Pt
from pytesseract import image_to_string
from PIL import Image

def setup_logger():
    try:
        # Ensure logs directory exists
        os.makedirs("logs", exist_ok=True)
        
        # Remove default handler
        logger.remove()
        
        # Configure handlers
        logger.add(
            sys.stdout,
            level="INFO",  # Changed from DEBUG to INFO for console
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>",
            colorize=True
        )

        # Add a file logger at DEBUG (more detailed logging for files)
        logger.add(
            "logs/applog_{time:YYYY-MM-DD}.log",  # Added date to filename
            level="DEBUG",
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name}:{function}:{line} - {message}",
            rotation="00:00",    # Rotate at midnight
            retention="30 days", # Increased retention period
            compression="zip",
            backtrace=True,     # Enable detailed traceback
            diagnose=True,      # Enable diagnostic mode
            enqueue=True        # Thread-safe logging
        )
        
        return logger
    except Exception as e:
        print(f"Failed to initialize logger: {e}")
        raise

def load_config(config_file: str) -> dict:
    """
    Loads configuration from a YAML file.
    """
    if not os.path.isfile(config_file):
        logging.error(f"Config file '{config_file}' not found.")
        sys.exit(1)

    try:
        with open(config_file, 'r', encoding='utf-8') as file:
            config = yaml.safe_load(file)
            if config is None:
                raise ValueError("Empty or invalid YAML structure.")
    except (yaml.YAMLError, ValueError) as e:
        logging.error(f"Failed to parse YAML config: {e}")
        sys.exit(1)

    return config

def validate_config(config: dict) -> None:
    """
    Validates essential keys and types in the configuration.
    Exits if required keys are missing or invalid.
    """
    if "paths" not in config:
        logging.error("Missing 'paths' section in config.yaml.")
        sys.exit(1)
    if "presentation" not in config:
        logging.error("Missing 'presentation' section in config.yaml.")
        sys.exit(1)

    paths = config["paths"]
    required_paths_keys = ["images_folder", "output_folder", "output_filename"]
    for key in required_paths_keys:
        if key not in paths:
            logging.error(f"Missing '{key}' under 'paths' in config.yaml.")
            sys.exit(1)

    images_folder = paths["images_folder"]
    output_folder = paths["output_folder"]
    output_filename = paths["output_filename"]

    if not os.path.isdir(images_folder):
        logging.error(f"Images folder '{images_folder}' does not exist or is not a directory.")
        sys.exit(1)

    if not isinstance(output_filename, str) or not output_filename.endswith(".pptx"):
        logging.error(f"Output filename '{output_filename}' must be a string ending with .pptx.")
        sys.exit(1)

    presentation = config["presentation"]
    required_presentation_keys = [
        "textbox_left_inches",
        "textbox_top_inches",
        "textbox_width_inches",
        "textbox_height_inches",
        "image_left_inches",
        "image_top_inches",
        "image_scale_percent",
        "text_font_size"
    ]

    for key in required_presentation_keys:
        if key not in presentation:
            logging.error(f"Missing '{key}' under 'presentation' in config.yaml.")
            sys.exit(1)
        try:
            float(presentation[key])
        except (TypeError, ValueError):
            logging.error(f"'{key}' under 'presentation' must be a numerical value.")
            sys.exit(1)

    # Validate 'extensions' if present
    if "extensions" in config:
        if not isinstance(config["extensions"], list):
            logging.error("'extensions' in config.yaml should be a list of file extensions.")
            sys.exit(1)
        for ext in config["extensions"]:
            if not (isinstance(ext, str) and ext.startswith(".")):
                logging.error(
                    f"Invalid extension '{ext}' in 'extensions'; must be a string like '.png'."
                )
                sys.exit(1)

def create_powerpoint_slides(config: dict) -> None:
    """
    Reads image files, performs OCR, and creates a PowerPoint presentation.
    """
    images_folder = config["paths"]["images_folder"]
    output_folder = config["paths"]["output_folder"]
    output_filename = config["paths"]["output_filename"]

    presentation_cfg = config["presentation"]

    textbox_left = float(presentation_cfg["textbox_left_inches"])
    textbox_top = float(presentation_cfg["textbox_top_inches"])
    textbox_width = float(presentation_cfg["textbox_width_inches"])
    textbox_height = float(presentation_cfg["textbox_height_inches"])

    image_left = float(presentation_cfg["image_left_inches"])
    image_top = float(presentation_cfg["image_top_inches"])
    image_scale_percent = float(presentation_cfg["image_scale_percent"])
    text_font_size = float(presentation_cfg["text_font_size"])

    slide_size_option = presentation_cfg.get("slide_size_option", "widescreen").lower()
    size_map = {
        "standard":   (10.0, 7.5),
        "widescreen": (13.3333, 7.5)
    }
    if slide_size_option not in size_map:
        logging.warning(f"Invalid slide_size_option '{slide_size_option}'. Defaulting to 'widescreen'.")
        slide_size_option = "widescreen"

    slide_width, slide_height = size_map[slide_size_option]
    allowed_extensions = [ext.lower() for ext in config.get("extensions", [".png"])]

    logging.info(f"Images folder: {images_folder}")
    logging.info(f"Output folder: {output_folder}")
    logging.info(f"Output filename: {output_filename}")
    logging.info(f"Allowed extensions: {allowed_extensions}")
    logging.info(f"Using slide size: {slide_width}in x {slide_height}in")

    os.makedirs(output_folder, exist_ok=True)
    full_output_path = os.path.join(output_folder, output_filename)

    presentation = Presentation()
    presentation.slide_width = Inches(slide_width)
    presentation.slide_height = Inches(slide_height)

    try:
        all_files = sorted(os.listdir(images_folder))
    except OSError as e:
        logging.error(f"Error reading images folder '{images_folder}': {e}")
        sys.exit(1)

    image_files = [
        f for f in all_files
        if any(f.lower().endswith(ext) for ext in allowed_extensions)
    ]

    if not image_files:
        logging.warning("No valid image files found in the specified folder.")

    for image_file in image_files:
        image_path = os.path.join(images_folder, image_file)
        logging.info(f"Processing image: {image_path}")

        try:
            with Image.open(image_path) as img:
                img.seek(0)  # handle multi-frame
                extracted_text = image_to_string(img)
                orig_width_px, orig_height_px = img.size
        except OSError as e:
            logging.error(f"Could not open or read image '{image_path}': {e}")
            continue

        slide = presentation.slides.add_slide(presentation.slide_layouts[6])

        # Convert px to inches based on 96 dpi assumption or fallback
        dpi_assumption = 96.0
        base_width_in = orig_width_px / dpi_assumption
        base_height_in = orig_height_px / dpi_assumption
        scale_factor = image_scale_percent / 100.0
        scaled_width_in = base_width_in * scale_factor
        scaled_height_in = base_height_in * scale_factor

        try:
            slide.shapes.add_picture(
                image_path,
                Inches(image_left),
                Inches(image_top),
                width=Inches(scaled_width_in),
                height=Inches(scaled_height_in)
            )
        except OSError as e:
            logging.error(f"Could not add image '{image_path}' to slide: {e}")
            continue

        text_box = slide.shapes.add_textbox(
            Inches(textbox_left),
            Inches(textbox_top),
            Inches(textbox_width),
            Inches(textbox_height)
        )
        text_frame = text_box.text_frame
        text_frame.text = extracted_text

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(text_font_size)

    try:
        presentation.save(full_output_path)
        logging.info(f"PowerPoint presentation saved to: {full_output_path}")
    except OSError as e:
        logging.error(f"Failed to save PowerPoint to '{full_output_path}': {e}")
        sys.exit(1)

def main() -> None:
    setup_logger()
    config = load_config("config.yaml")
    validate_config(config)
    create_powerpoint_slides(config)

if __name__ == "__main__":
    main()