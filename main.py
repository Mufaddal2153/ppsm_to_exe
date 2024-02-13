import os
import subprocess
import sys
import tempfile
import zipfile
import shutil

def extract_zip_to_temp(zip_path):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir

def run_ppt(temp, ppsm_file_path):
    pptview_exe_path = os.path.join(temp, "Office14/PPTVIEW.EXE")

    if not os.path.isfile(pptview_exe_path):
        print(f"Error: PPTVIEW.EXE not found.")
        sys.exit(1)

    command = [pptview_exe_path, "/s", "/f", ppsm_file_path]

    try:
        subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")
        sys.exit(1)

if __name__ == "__main__":
    bundled_zip_path = "Office14.zip"
    zip_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), bundled_zip_path)
    temp_dir = extract_zip_to_temp(zip_path)
    ppsm_file_path = os.path.join(temp_dir, "Office14/etisalat by e&.ppsx")
    run_ppt(temp_dir, ppsm_file_path)
    shutil.rmtree(temp_dir)
