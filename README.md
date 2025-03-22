images2pptx
===========

Description
-----------
This project automates:

1.  **Reading** images from a folder.
2.  **Extracting text** from them via **OCR** (using [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)).
3.  **Generating** a PowerPoint presentation, placing the images and extracted text on each slide.

## Give a Star! ⭐

If you enjoy this project, please consider giving it a star—it helps others discover it too!

* * *

Table of Contents
-----------------

1.  [Requirements](#requirements)
2.  [Setup Instructions](#setup-instructions)
    1.  [Install Python (if needed)](#install-python-if-needed)
    2.  [Create a Virtual Environment](#create-a-virtual-environment)
    3.  [Activate the Virtual Environment](#activate-the-virtual-environment)
    4.  [Install Dependencies](#install-dependencies)
3.  [Configuration](#configuration)
    1.  [Sample `config.yaml` Explanation](#sample-configyaml-explanation)
    2.  [Notes About the Paths](#notes-about-the-paths)
4.  [Usage](#usage)
5.  [Troubleshooting / Common Issues](#troubleshooting--common-issues)
6.  [License](#license)

* * *

Requirements
------------

*   **Python** 3.8+ (earlier versions may work, but are untested)
*   **Tesseract OCR** command-line tool (so `pytesseract` can interface with it).
    *   On macOS: `brew install tesseract`
    *   On Ubuntu/Debian: `sudo apt-get install tesseract-ocr`
    *   On Windows: [Download installer from tesseract\-ocr/tesseract GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
*   A modern operating system (Windows, macOS, Linux) with **pip** available.

* * *

Setup Instructions
------------------

### 1\. Install Python (if needed)

If you do not already have a compatible Python version (3.8 or above), install it from:

*   [https://www.python.org/downloads/](https://www.python.org/downloads/)

Make sure to check the box to **Add Python to PATH** on Windows.

### 2\. Create a Virtual Environment

It’s a best practice to run this project within a **virtual environment** (venv). Inside your project folder (where `main.py` and `config.yaml` are located), open a terminal (or command prompt) and run:

```bash
# On macOS/Linux:
python3 -m venv venv

# On Windows:
python -m venv venv
```

This creates a folder named `venv` that contains an isolated Python environment.

### 3\. Activate the Virtual Environment

*   **macOS/Linux**:
    
    ```bash
    source venv/bin/activate
    ```
    
*   **Windows (Command Prompt)**:
    
    ```bash
    venv\Scripts\activate
    ```
    
*   **Windows (PowerShell)**:
    
    ```powershell
    .\venv\Scripts\activate.ps1
    ```
    

You should notice your prompt change to indicate you’re now inside the `venv` environment.

### 4\. Install Dependencies

Inside your virtual environment, install the necessary packages:

```bash
pip install -r requirements.txt
```

If you don’t have a `requirements.txt`, you can manually install:

```bash
pip install python-pptx pytesseract Pillow pyyaml
```

> **Note**: Make sure **Tesseract** is installed on your system so that `pytesseract` can run OCR.

* * *

Configuration
-------------

### 1\. Sample `config.yaml` Explanation

A typical `config.yaml` can look like this:

```yaml
paths:
  images_folder: "/path/to/your/images_folder"   # The folder containing the images you want to process
  output_folder: "/path/to/your/output_folder"   # Where the resulting .pptx will be saved
  output_filename: "output_presentation.pptx"    # The name of the generated PowerPoint

presentation:
  # Slide size: either "standard" (4:3) or "widescreen" (16:9)
  # Defaults to "widescreen" if omitted or invalid
  slide_size_option: "widescreen"

  # Textbox placement and size in inches
  textbox_left_inches: 1
  textbox_top_inches: 1
  textbox_width_inches: 3
  textbox_height_inches: 5

  # Image placement and scale
  image_left_inches: 6
  image_top_inches: 0.5
  image_scale_percent: 50   # Example: 50 -> 50% scale
  text_font_size: 10        # Font size for OCR text

extensions:
  - ".png"
  - ".jpg"
  - ".jpeg"
  - ".gif"
```

#### Explanation of Key Fields

1.  **`paths.images_folder`**: The directory containing your input images.
    
2.  **`paths.output_folder`**: The directory where you want the `.pptx` file to be saved.
    
3.  **`paths.output_filename`**: The name for the PowerPoint file (must end in `.pptx`).
    
4.  **`presentation.slide_size_option`**: Chooses between:
    
    *   **`standard`**: Creates a 4:3 slide size (10 in x 7.5 in).
    *   **`widescreen`**: Creates a 16:9 slide size (13.3333 in x 7.5 in).
5.  **`presentation.textbox_left_inches`, `presentation.textbox_top_inches`, etc.**: Coordinates (inches) within the slide for placing your text box.
    
6.  **`presentation.image_left_inches` & `image_top_inches`**: Coordinates (inches) within the slide for placing the **top-left corner** of each image.
    
7.  **`presentation.image_scale_percent`**: A numeric scale factor for images (e.g., 50 means 50% of original size).
    
8.  **`presentation.text_font_size`**: Size (in points) for the text that will be inserted from OCR.
    
9.  **`extensions`**: A list of file extensions that should be treated as images. Only these files in `images_folder` will be processed.
    

### 2\. Notes About the Paths

*   Use **forward slashes** `/` or **escaped** backslashes `\\` in Windows.
*   Ensure the path values are **quoted** if they contain special characters or spaces.
*   If the path doesn’t exist for `output_folder`, the script will create it for you.

* * *

Usage
-----

1.  **Activate** your virtual environment (if not already active).
2.  **Make sure** your `config.yaml` is properly set up (or you can rename a provided `config.example.yaml` to `config.yaml` and adjust the paths as needed).
3.  **Run** the main script:
    
    ```bash
    python main.py
    ```
    
4.  Watch your terminal for logs. OCR processing may take a bit of time if you have many images.
5.  A file called (for example) `output_presentation.pptx` should appear in the `output_folder` directory.

* * *

Troubleshooting / Common Issues
-------------------------------

1.  **Tesseract Not Found**
    
    *   Ensure Tesseract is installed and available on your system path. On some systems, you may need to specify the Tesseract executable path in code or in an environment variable. See [pytesseract documentation](https://pypi.org/project/pytesseract/) for details.
2.  **Pillow / PIL Errors**
    
    *   Some images with unusual color profiles or corrupt metadata can cause errors. Check that your images open in standard image viewers.
3.  **No Slides Created**
    
    *   Make sure the `extensions` list in `config.yaml` matches your image file types. Check that `images_folder` is correct and not empty (or containing subfolders you may not have accounted for).
4.  **No text extracted**
    
    *   OCR can fail if the images have very small text, are low contrast, or use unusual fonts. Check that Tesseract is functioning by running `tesseract <image> stdout` in your terminal.
5.  **File Permissions**
    
    *   On macOS/Linux, ensure you have the right permissions to read from `images_folder` and write to `output_folder`.

* * *

License
-------

This repository is published under the MIT License (or your license of choice). See LICENSE for details. Feel free to modify this as needed for your specific license situation.

* * *

That’s it! With these steps, you and other users should be able to set up a Python environment, configure `config.yaml`, and run the script to generate PowerPoint files from images via OCR. 