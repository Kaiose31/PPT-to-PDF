from itertools import repeat
import sys
import os
import win32com.client
import pythoncom
from multiprocessing import Pool, freeze_support

def paraconv(input_file_name,input_folder_path):
    pythoncom.CoInitialize()
    if not input_file_name.lower().endswith((".pptx", ".pptm")):
        return
    pptx_path  =  os.path.abspath(os.path.join(input_folder_path,input_file_name))
    output_path =  pptx_path.split('.')[0]+'.pdf'
    output_path  =  os.path.abspath(output_path)
    powerpoint = win32com.client.DispatchEx("Powerpoint.Application")
    slides = powerpoint.Presentations.Open(pptx_path,WithWindow=False)
    slides.SaveAs(output_path, 32)
    slides.Close()
    return output_path

def convertall(input_folder_path):
    input_files = os.listdir(input_folder_path)
    pool = Pool()
    pool.starmap(paraconv,zip(input_files,repeat(input_folder_path)))
    os.system("TASKKILL /F /IM powerpnt.exe")

if __name__ == "__main__":
    freeze_support()
    convertall(sys.argv[1])
