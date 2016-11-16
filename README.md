# metadatos
Script that recovers metadata from PDF, DOCX, JPEG, PNG, GIF, BMP and TIFF files. 

Module dependencies (installable via pip): 
* PyPDF2
* docx
* PIL (pillow)
* exifread
* libxmp (python-xmp-toolkit)
* termcolor

## How to use
You just have to execute the script with the folder containing the files to be analysed as an argument. 

For instance, this will analyse all files inside ./samples. 
```
python metadatos.py ./samples
```
