import os
import shutil
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import shutil
import zipfile

from backend.services.excel_generator_v4 import process_packing_slip as process_v4
from backend.services.pdf_service import convert_excel_to_pdf, split_pdf
from backend.services.test_readings_service import get_test_readings, update_test_reading

app = FastAPI()

# Allow frontend to communicate
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # In production, restrict to frontend URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "backend_output")

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app.mount("/outputs", StaticFiles(directory=OUTPUT_DIR), name="outputs")

@app.post("/upload")
async def upload_packing_slip(file: UploadFile = File(...), output_mode: str = Form("single")):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported")
        
    try:
        import time
        run_id = str(int(time.time()))
        run_output_dir = os.path.join(OUTPUT_DIR, run_id)
        os.makedirs(run_output_dir, exist_ok=True)
        
        input_path = os.path.join(UPLOADS_DIR, f"{run_id}_{file.filename}")
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        if output_mode == "group":
            master_excel_path, num_pages, group_mapping = process_v4(input_path, run_output_dir, output_mode)
            
            output_pdf_path = os.path.join(run_output_dir, "Master_All_Reports.pdf")
            try:
                convert_excel_to_pdf(master_excel_path, output_pdf_path)
                split_pdf(output_pdf_path, group_mapping, run_output_dir)
            except Exception as pdf_err:
                print(f"CloudConvert/PDF generation failed, falling back to Excels only: {pdf_err}")
            
            # Zip the directory (will contain PDFs if successful, or just Excels if failed)
            zip_filename = f"Group_Reports_{run_id}.zip"
            zip_path = os.path.join(OUTPUT_DIR, zip_filename)
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(run_output_dir):
                    for f in files:
                        file_path = os.path.join(root, f)
                        arcname = os.path.relpath(file_path, run_output_dir)
                        zipf.write(file_path, arcname)
                        
            return JSONResponse(content={
                "file_url": f"/outputs/{zip_filename}",
                "total_pages": num_pages,
                "type": "zip"
            })
            
        else:
            master_excel_path, num_pages, _ = process_v4(input_path, run_output_dir, output_mode)
            
            output_pdf_path = os.path.join(run_output_dir, f"Final_Inspection_{run_id}.pdf")
            try:
                convert_excel_to_pdf(master_excel_path, output_pdf_path)
                final_hosted_file = output_pdf_path
                file_type = "pdf"
            except Exception as pdf_err:
                print(f"CloudConvert/PDF generation failed, falling back to Excel only: {pdf_err}")
                # Fallback to the master excel file
                final_hosted_file = os.path.join(run_output_dir, f"Master_All_Reports_{run_id}.xlsx")
                shutil.copy2(master_excel_path, final_hosted_file)
                file_type = "excel"
            
            # Move the final output to the accessible public dir
            public_filename = os.path.basename(final_hosted_file)
            public_filepath = os.path.join(OUTPUT_DIR, public_filename)
            shutil.copy2(final_hosted_file, public_filepath)
            
            return JSONResponse(content={
                "file_url": f"/outputs/{public_filename}",
                "total_pages": num_pages,
                "type": file_type
            })
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/test-readings")
async def fetch_test_readings():
    try:
        data = get_test_readings()
        return JSONResponse(content=data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

from pydantic import BaseModel
from typing import List, Any

class UpdatePayload(BaseModel):
    row: int
    col: int
    val: Any

@app.post("/api/test-readings/update")
async def save_test_readings(updates: List[UpdatePayload]):
    try:
        update_list = [{"row": u.row, "col": u.col, "val": u.val} for u in updates]
        update_test_reading(update_list)
        return JSONResponse(content={"message": "Successfully updated readings"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/upload-tracker")
async def upload_tracker(file: UploadFile = File(...)):
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported")
    
    try:
        tracker_path = os.path.join(BASE_DIR, "sample_inputs", "Item Codes Tracker.xlsx")
        os.makedirs(os.path.dirname(tracker_path), exist_ok=True)
        
        with open(tracker_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        return JSONResponse(content={"message": "Tracker successfully updated!"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
