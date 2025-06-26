import os
import sys
from pathlib import Path
import win32com.client

def convert_ppt_to_pptx(ppt_path, pptx_path):
    try:
        powerpoint = win32com.client.Dispatch('PowerPoint.Application')
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(str(ppt_path), WithWindow=False)
        presentation.SaveAs(str(pptx_path), FileFormat=24)  # 24: pptx
        presentation.Close()
        powerpoint.Quit()
        print(f"Converted: {ppt_path} -> {pptx_path}")
    except Exception as e:
        print(f"Error occurred: {ppt_path}: {e}")

def main(ppt_folder):
    ppt_folder = Path(ppt_folder)
    if not ppt_folder.exists() or not ppt_folder.is_dir():
        print("Please enter a valid folder path.")
        return
    ppt_files = list(ppt_folder.glob('*.ppt'))
    if not ppt_files:
        print("No PPT files found in the folder.")
        return
    for ppt_file in ppt_files:
        pptx_file = ppt_file.with_suffix('.pptx')
        convert_ppt_to_pptx(ppt_file, pptx_file)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python ppt_to_pptx.py <ppt_folder>")
    else:
        main(sys.argv[1]) 