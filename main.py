import os
import subprocess
import sys
import tempfile
import zipfile
import shutil

def execute_exe(exe_path):
    if not os.path.isfile(exe_path):
        print(f"Error: {exe_path} not found.")
        sys.exit(1)
    try:
        subprocess.call(exe_path, shell=True)
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {e}")
        sys.exit(1)

def extract_zip_to_temp(zip_path):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir

def run_ppt(temp, ppsm_file_path):
    pptview_exe_path = os.path.join(temp, "PPTVIEW.EXE")

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
    bundled_zip_path = "C:\Program Files (x86)\Microsoft Office\Office14"
    zip_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "pptview.zip")
    ezip_path = extract_zip_to_temp(zip_path)
    if not os.path.isdir(bundled_zip_path):
        execute_exe(os.path.join(ezip_path, "powerpoint-viewer_softradar-com.exe"))
    
    ppsm_file_path = os.path.join(ezip_path, "e& Life.ppsx")
    run_ppt(bundled_zip_path, ppsm_file_path)
    shutil.rmtree(ezip_path)
