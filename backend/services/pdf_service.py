import os
import urllib.request
import cloudconvert
from pypdf import PdfReader, PdfWriter

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

def convert_excel_to_pdf(excel_path, output_pdf_path):
    # Read API Key
    api_key_path = os.path.join(BASE_DIR, 'api_key.txt')
    with open(api_key_path, 'r') as f:
        api_key = f.read().strip()

    cloudconvert.configure(api_key=api_key)

    job = cloudconvert.Job.create(payload={
        'tasks': {
            'import-my-file': {
                'operation': 'import/upload'
            },
            'convert-my-file': {
                'operation': 'convert',
                'input': 'import-my-file',
                'output_format': 'pdf',
                'engine': 'libreoffice' 
            },
            'export-my-file': {
                'operation': 'export/url',
                'input': 'convert-my-file'
            }
        }
    })

    upload_task = next(task for task in job['tasks'] if task['name'] == 'import-my-file')
    cloudconvert.Task.upload(file_name=excel_path, task=upload_task)

    # Wait for processing
    job = cloudconvert.Job.wait(id=job['id'])

    export_task = next(task for task in job['tasks'] if task['name'] == 'export-my-file')

    if export_task['status'] == 'finished':
        file_url = export_task['result']['files'][0]['url']
        urllib.request.urlretrieve(file_url, output_pdf_path)
        return output_pdf_path
    else:
        raise Exception(f"CloudConvert failed: {export_task}")

def split_pdf(master_pdf_path, group_mapping, output_dir):
    """
    Slices the master PDF into individual Group PDFs based on page index mapping.
    group_mapping: dict { "Group Name": (start_page_idx, end_page_idx) }
    output_dir: Base directory where Group folders exist
    """
    reader = PdfReader(master_pdf_path)
    
    for group_name, (start_idx, end_idx) in group_mapping.items():
        writer = PdfWriter()
        # Add pages for this group
        for i in range(start_idx, end_idx + 1):
            if i < len(reader.pages):
                writer.add_page(reader.pages[i])
                
        safe_group = group_name.replace("/", "_").replace(":", "_").strip()
        group_folder = os.path.join(output_dir, safe_group)
        os.makedirs(group_folder, exist_ok=True)
        
        pdf_path = os.path.join(group_folder, f"Master_{safe_group}.pdf")
        with open(pdf_path, "wb") as f:
            writer.write(f)
            
    return True

