# Quality Inspection Automation - Project Context & Architecture

This document provides a high-level overview of the current features, data flow, and architecture of the Quality Inspection Automation system.

## 1. Current Features

* **Automated Packing Slip Processing:** Users can upload a standard Excel Packing Slip via the drag-and-drop web interface.
* **Intelligent Data Cross-Referencing:** The system automatically cross-references item codes from the Packing Slip with the Master Item Codes Tracker to determine exact specifications (Type, Copper Type, Nominal OD, Colour, etc.).
* **Dynamic Test Readings Generation:** Generates realistic, within-spec randomized test readings based on base profiles found in `TEST READINGS.xlsx`.
* **Pixel-Perfect Excel Generation:** Programmatically builds highly formatted, multi-page inspection reports in Excel using `openpyxl`, complete with company logos, approval stamps, precision cell borders, and specific fonts.
* **Automated PDF Conversion:** Seamlessly merges all individual reports into a Master Excel file and converts it into a single, print-ready PDF using the CloudConvert API.
* **Modern Web Interface:** A React-based frontend that provides real-time status updates and an embedded interactive PDF preview of the final generated report.

## 2. Architecture & Data Flow

The system is split into a decoupled **Frontend (React/Vite)** and **Backend (FastAPI/Python)**.

### Step 1: Upload (Frontend -> Backend)
1. The user drags and drops a `Packing Slip.xlsx` file onto the React frontend (running on `localhost:5175`).
2. The frontend sends this file via a `POST /upload` request to the FastAPI backend (running on `localhost:8090`).
3. The backend saves the uploaded file into the local `uploads/` directory for processing.

### Step 2: Data Parsing & Grouping (`excel_generator.py`)
1. **Packing Slip Parsing:** The backend uses `pandas` to read the uploaded Packing Slip. It scans for valid item codes and groups the data by **Invoice Number** and **Item Code**, summing up the total length and total number of coils.
2. **Item Codes Tracker Lookups:** For every unique item code, the system reads `sample_inputs/Item Codes Tracker.xlsx` to extract specifications like `Type`, `Copper Type`, `Nominal OD`, and `Colour`. 
3. **Group Name Normalization:** The system builds a constructed name (e.g., `AVX 0.5F OD 1.70`) and runs it through a `normalize_loose()` function (which strips special characters, spaces, and strings like 'mm' and 'f'). This normalized string is matched against `group_names/Group Names.xlsx` to find the official Group Name classification.

### Step 3: Fetching Test Readings
1. The system reads `sample_inputs/TEST READINGS.xlsx` to find a "Base Profile" that matches the item's `Copper Type` and `Nominal OD`.
2. This base profile provides the baseline measurements for critical parameters like:
   * Strands Dia
   * Insulation Dia
   * Conductor Resistance
   * Insulation Thickness (Nom.)
3. The generator then applies slight, controlled random variations to these base readings to generate 8 distinct, realistic observation columns per coil (to simulate actual human testing).

### Step 4: Excel Document Generation
1. Using `openpyxl`, the system generates a meticulously formatted Excel worksheet for every unique item.
2. **Formatting Logic:** 
   * Sets page orientation to A4 Landscape with `fitToPage = True`.
   * Inserts the Elentec and Vardhan logos at the top (`Row 1-2`).
   * Draws dynamic outer borders strictly around the merged header cells to create clean boxes.
   * Populates the 8 generated test readings into the observation columns.
   * Injects the digital approval stamp at the bottom of the sheet, scaled perfectly to fit.
3. All individual worksheets are combined and saved as a single `Master_All_Reports.xlsx` file in the `v2output/` (or `backend_output/`) directory.

### Step 5: PDF Conversion & Delivery
1. **CloudConvert Integration:** The backend uses the CloudConvert API (`pdf_service.py`) to upload the `Master_All_Reports.xlsx` file, convert it to PDF format, and download the resulting `Final_Inspection_Report.pdf`.
2. **Static Serving:** The FastAPI backend serves the `backend_output/` folder statically. 
3. **Frontend Render:** The backend responds to the initial upload request with a JSON object containing the static URL of the generated PDF. The React frontend receives this URL and embeds it into an `<iframe>` for instant preview and provides a direct download button.

## Directory Structure Overview
* **`frontend/`**: Contains the React/Vite web application.
* **`backend/`**: Contains the FastAPI server, endpoints, and CloudConvert integration logic.
* **`backend/services/excel_generator.py`**: The core engine for parsing Excel logic and drawing the openpyxl worksheets.
* **`sample_inputs/`**: Contains the reference files needed by the backend (`Item Codes Tracker.xlsx`, `TEST READINGS.xlsx`, etc.).
* **`backend_output/` / `v2output/`**: The directories where the generated Excel sheets, Master Excel, and Final PDF are stored.
