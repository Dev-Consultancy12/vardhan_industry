import os
import shutil
import zipfile
from backend.services.excel_generator_v4 import process_packing_slip

BASE_DIR = "/Users/namanbansal/vardhan_industry_automation"
INPUT_SLIP = os.path.join(BASE_DIR, "uploads/Packing Slip 28-31.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "v3output")

def run_manual():
    print("Running Group-Based V4...")
    group_out = os.path.join(OUTPUT_DIR, "group_based_28_31")
    os.makedirs(group_out, exist_ok=True)
    process_packing_slip(INPUT_SLIP, group_out, "group")
    
    zip_path = os.path.join(OUTPUT_DIR, "Group_Reports_28_31_Excels.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(group_out):
            for f in files:
                file_path = os.path.join(root, f)
                arcname = os.path.relpath(file_path, group_out)
                zipf.write(file_path, arcname)
    print(f"Saved to {zip_path}")

    print("Running Single V4...")
    single_out = os.path.join(OUTPUT_DIR, "single_28_31")
    os.makedirs(single_out, exist_ok=True)
    process_packing_slip(INPUT_SLIP, single_out, "single")
    
    single_zip = os.path.join(OUTPUT_DIR, "Single_Reports_28_31_Excels.zip")
    with zipfile.ZipFile(single_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(single_out):
            for f in files:
                file_path = os.path.join(root, f)
                arcname = os.path.relpath(file_path, single_out)
                zipf.write(file_path, arcname)
    print(f"Saved to {single_zip}")

if __name__ == "__main__":
    run_manual()
