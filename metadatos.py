#!/usr/bin/python
# -*- encoding: utf-8 -*- 

from PyPDF2 import PdfFileReader,PdfFileWriter # Module for PDF manipulation
import docx # Module for DOCX manipulation
from PIL import Image # Module for image manipulation
from PIL.ExifTags import TAGS, GPSTAGS
import exifread # Module for EXIF metadata manipulation
from libxmp.utils import file_to_dict # Module for XMP metadata manipulation
from libxmp import consts
import os # Module for OS interaction
import sys # Module for system interaction
from termcolor import cprint # Module for terminal colored output
import datetime

"""
CREATED BY GUILLERMO ROMAN FERRERO a.k.a. hartek
This script prints the metadata of the files found recursively in a directory given by argument. 
Currently supports PDF, DOCX, JPEG, PNG, GIF, BMP and TIFF files. 
"""

def printMeta(target): 
	"""Iterates the target directory in search for documents"""
	if not os.path.isdir(target): 
		cprint("Cound not find target directory", "red")
		return
	walk = os.walk(target)
	for dirpath, dirnames, files in walk: 
		for name in files: 
			ext = name.lower().rsplit(".", 1)[-1]
			file_full_path = dirpath+os.path.sep+name
			if ext == "pdf":
				print_pdf(file_full_path)
			elif ext == "docx":
				print_docx(file_full_path)
			elif ext in ("jpg","jpeg",): 
				print_jpg(file_full_path)
			elif ext in ("png","gif","bmp"): 
				print_png_gif_bmp(file_full_path)
			elif ext in ("tif","tiff"):
				print_tiff(file_full_path)
			else: 
				pass


def print_jpg(file_full_path): 
	""" Analyzes the metadate of a JPG/JPEG file """
	# Header with file path
	cprint("[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	# Open the file
	image = Image.open(file_full_path)
	# Print XMP metadata
	cprint("\t-----XMP METADATA-----", "cyan")
	xmp = file_to_dict(file_full_path)
	if not xmp: 
		cprint("\tNo XMP metadata found", "red")
	else: 
		dc = xmp[consts.XMP_NS_DC]
		cprint("\t-" + dc[0][0], "cyan")
		cprint("\t-" + dc[0][1], "cyan")

		for key, value in dc[0][2].items(): 
			cprint("\t-" + key + ": ", "cyan", end="")
			cprint(str(value))
	# Print EXIF metadata
	cprint("\t-----EXIF METADATA-----", "cyan")
	info = image._getexif()
	if not info: 
		cprint("\tNo XMP metadata found", "red")
	else:
		for tag, value in info.items():
		    key = TAGS.get(tag, tag)
		    print(key + " " + str(value))

def print_png_gif_bmp(file_full_path): 
	""" Analyzes the metadate of a PNG file """
	# Header with file path
	cprint("[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])	
	# Open the file
	image = Image.open(file_full_path)
	
	# Print XMP metadata
	cprint("\t-----XMP METADATA-----", "cyan")
	xmp = file_to_dict(file_full_path)
	if not xmp: 
		cprint("\tNo XMP metadata found", "red")
	else: 
		dc = xmp[consts.XMP_NS_DC]
		cprint("\t-" + dc[0][0], "cyan")
		cprint("\t-" + dc[0][1], "cyan")

		for key, value in dc[0][2].items(): 
			cprint("\t-" + key + ": ", "cyan", end="")
			cprint(str(value))


def print_tiff(file_full_path): 
	""" Analyzes the metadate of a JPG/JPEG file """
	# Header with file path
	cprint("[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	# Open the file
	image = open(file_full_path, 'rb')

	# Print XMP metadata
	cprint("\t-----XMP METADATA-----", "cyan")
	xmp = file_to_dict(file_full_path)
	dc = xmp[consts.XMP_NS_DC]
	cprint("\t-" + dc[0][0], "cyan")
	cprint("\t-" + dc[0][1], "cyan")

	for key, value in dc[0][2].items(): 
		cprint("\t-" + key + ": ", "cyan", end="")
		cprint(str(value))

	# Print EXIF metadata
	cprint("\n\t-----EXIF METADATA-----", "cyan")
	tags = exifread.process_file(image)
	for tag in tags.keys():
	    if tag not in ('JPEGThumbnail', 'TIFFThumbnail', 'Filename', 'EXIF MakerNote'):
	    	cprint("\t-" + str(tag) + ": ", "cyan", end="")
	    	cprint(str(tags[tag]))


def print_docx(file_full_path): 
	"""Analyzes the metadata of a .docx file"""
	# Header with file path
	cprint("[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	# Open the file
	docxFile = docx.Document(file(file_full_path, "rb"))
	# Data structure with document information
	docxInfo = docxFile.core_properties
	# Print metadata
	attrs = ["author", "category", "comments", "content_status", 
		"created", "identifier", "keywords", "language", 
		"last_modified_by", "last_printed", "modified", 
		"revision", "subject", "title", "version"]
	for attr in attrs: 
		value = getattr(docxInfo,attr)
		if value:
			if isinstance(value, unicode): 
				cprint("\t-" + str(attr) + ": ", "cyan", end="")
				cprint(value)
			elif isinstance(value, datetime.datetime):  
				cprint("\t-" + str(attr) + ": ", "cyan", end="")
				cprint(str(value))
	print ""

def print_pdf(file_full_path):
	"""Analyzes the metadata of a .pdf file"""
	# Header with file path
	cprint("[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	# Open the file
	pdf_file = PdfFileReader(file(file_full_path, "rb"))
	# Data structure with document information
	pdf_info = pdf_file.getDocumentInfo()
	# Print metadata
	if pdf_info: 
		for metaItem in pdf_info: 
			try: 
				cprint("\t-" + metaItem[1:] + ": ", "cyan", end="")
				cprint(pdf_info[metaItem])
			except TypeError: 
				cprint("\t-" + metaItem[1:] + ": " + "Error - Item not readable", "red")
	else: 
		cprint("\t No data found", "red")
	print ""
# Main function
def main (argv):
	"""Main function of the program"""
	# Check arguments
	if len(argv) != 2: 
		print "Incorrect number of arguments. "
		return
	# If arguments OK, execute function
	else: 
		target = argv[1]
		printMeta(target)

# Execute main function
if __name__ == "__main__": 
	main(sys.argv)
