## Fast PowerPoint to PDF Converter
Uses all threads on the CPU to convert all files in a directory.
Compatible with PPTX and PPTM Format

* Requires Microsoft Powerpoint to be installed on the system.  (License not required).


### Requirements

1. Create and Activate the Virtual Environment
```
python -m venv  environment_name

environment_name\Scripts\activate.bat

```
2. install requirements.txt 
```
pip install -r requirements.txt
```
3. run pywin32_postinstall.py script
```
python environment_name\Scripts\pywin32_postinstall.py -install
```
4. Copy pythoncom38.dll and pywintypes38.dll from `environment_name\Lib\site-packages\pwin32_system32` to `environment_name\Lib\site-packages\win32\lib`

5. Run the script with targer folder as command line arguement.
```
python ppt_to_pdf.py target
```


