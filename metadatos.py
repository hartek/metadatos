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
import argparse
import datetime

"""
CREATED BY GUILLERMO ROMAN FERRERO a.k.a. hartek
This script prints the metadata of the files found recursively in a directory given by argument. 
Currently supports PDF, DOCX, JPEG, PNG, GIF, BMP and TIFF files. 
"""

def printMeta(target, color_mode): 
	"""Iterates the target directory in search for documents"""
	if color_mode: cprint("Trying to iterate directory...", "green", attrs=["bold"])
	else: print "Trying to iterate directory..."

	if not os.path.isdir(target): 
		if color_mode: ("Cound not find target directory", "red")
		else: print "Cound not find target directory"
		return
	walk = os.walk(target)
	for dirpath, dirnames, files in walk: 
		for name in files: 
			ext = name.lower().rsplit(".", 1)[-1]
			if dirpath[-1:] == "/": 
				file_full_path = dirpath+name
			else: 
				file_full_path = dirpath+os.path.sep+name
			if ext == "pdf":
				print_pdf(file_full_path, color_mode)
			elif ext == "docx":
				print_docx(file_full_path, color_mode)
			elif ext in ("jpg","jpeg",): 
				print_jpg(file_full_path, color_mode)
			elif ext in ("png","gif","bmp"): 
				print_png_gif_bmp(file_full_path, color_mode)
			elif ext in ("tif","tiff"):
				print_tiff(file_full_path, color_mode)
			else: 
				pass


def print_jpg(file_full_path, color_mode): 
	""" Analyzes the metadate of a JPG/JPEG file """
	# Header with file path
	if color_mode: 	cprint("\n[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	else: print "\n[+] Metadata for file: %s" % (file_full_path)
	# Open the file
	image = Image.open(file_full_path)
	# Print XMP metadata
	if color_mode: cprint("\t-----XMP METADATA-----", "cyan")
	else: print "\t-----XMP METADATA-----"

	xmp = file_to_dict(file_full_path)
	if not xmp: 
		if color_mode: cprint("\tNo XMP metadata found", "red")
		else: print "\tNo XMP metadata found"
	else: 
		dc = xmp[consts.XMP_NS_DC]
		if color_mode: 
			cprint("\t-" + dc[0][0], "cyan")
			cprint("\t-" + dc[0][1], "cyan")
		else: 
			print "\t-" + dc[0][0]
			print "\t-" + dc[0][1]

		for key, value in dc[0][2].items(): 
			if color_mode: 
				cprint("\t-" + key + ": ", "cyan", end="")
				cprint(str(value))
			else: 
				print "\t-" + key + ": "
				print str(value)
	# Print EXIF metadata
	if color_mode: cprint("\n\t-----EXIF METADATA-----", "cyan")
	else: print "\n\t-----EXIF METADATA-----"

	info = image._getexif()
	if not info: 
		if color_mode: cprint("\tNo EXIF metadata found", "red")
		else: print "\tNo EXIF metadata found"
	else:
		for tag, value in info.items():
		    key = TAGS.get(tag, tag)
		    if color_mode: 
		    	cprint("\t-" + key + ": ", "cyan", end="")
		    	cprint(str(value))
		    else: 
		    	print "\t-" + key + ": " + str(value)


def print_png_gif_bmp(file_full_path, color_mode): 
	""" Analyzes the metadate of a PNG file """
	# Header with file path
	if color_mode: cprint("\n[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])	
	else: print "\n[+] Metadata for file: %s" % (file_full_path)
	# Open the file
	image = Image.open(file_full_path)
	
	# Print XMP metadata
	if color_mode: cprint("\t-----XMP METADATA-----", "cyan")
	else: print "\t-----XMP METADATA-----"

	xmp = file_to_dict(file_full_path)
	if not xmp: 
		if color_mode: cprint("\tNo XMP metadata found", "red")
		else: print "\tNo XMP metadata found"
	else: 
		dc = xmp[consts.XMP_NS_DC]
		if color_mode: 
			cprint("\t-" + dc[0][0], "cyan")
			cprint("\t-" + dc[0][1], "cyan")
		else: 
			print "\t-" + dc[0][0]
			print "\t-" + dc[0][1]

		for key, value in dc[0][2].items(): 
			if color_mode: 
				cprint("\t-" + key + ": ", "cyan", end="")
				cprint(str(value))
			else: 
				print "\t-" + key + ": " + str(value)


def print_tiff(file_full_path, color_mode): 
	""" Analyzes the metadate of a JPG/JPEG file """
	# Header with file path
	if color_mode: cprint("\n[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	else: print "\n[+] Metadata for file: %s" % (file_full_path)
	# Open the file
	image = open(file_full_path, 'rb')

	# Print XMP metadata
	if color_mode: cprint("\t-----XMP METADATA-----", "cyan")
	else: "\t-----XMP METADATA-----"

	xmp = file_to_dict(file_full_path)
	if not xmp: 
		if color_mode: cprint("\tNo XMP metadata found", "red")
		else: "\tNo XMP metadata found"
	else: 
		dc = xmp[consts.XMP_NS_DC]
		if color_mode: 
			cprint("\t-" + dc[0][0], "cyan")
			cprint("\t-" + dc[0][1], "cyan")
		else: 
			print "\t-" + dc[0][0]
			print "\t-" + dc[0][1]

		for key, value in dc[0][2].items(): 
			if color_mode: 
				cprint("\t-" + key + ": ", "cyan", end="")
				cprint(str(value))
			else: 
				print "\t-" + key + ": " + str(value)

	# Print EXIF metadata
	if color_mode: print("\n\t-----EXIF METADATA-----", "cyan")
	else: "\n\t-----EXIF METADATA-----"

	tags = exifread.process_file(image)
	for tag in tags.keys():
	    if tag not in ('JPEGThumbnail', 'TIFFThumbnail', 'Filename', 'EXIF MakerNote'):
	    	if color_mode: 
	    		cprint("\t-" + str(tag) + ": ", "cyan", end="")
	    		cprint(str(tags[tag]))
	    	else: 
	    		print "\t-" + str(tag) + ": " + str(tags[tag])


def print_docx(file_full_path, color_mode): 
	"""Analyzes the metadata of a .docx file"""
	# Header with file path
	if color_mode: cprint("\n[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	else: "\n[+] Metadata for file: %s" % (file_full_path)
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
				if color_mode: 
					cprint("\t-" + str(attr) + ": ", "cyan", end="")
					cprint(value)
				else: 
					print "\t-" + str(attr) + ": " + value
			elif isinstance(value, datetime.datetime):  
				if color_mode: 
					cprint("\t-" + str(attr) + ": ", "cyan", end="")
					cprint(str(value))
				else: 
					print "\t-" + str(attr) + ": " + str(value)
	print ""

def print_pdf(file_full_path, color_mode):
	"""Analyzes the metadata of a .pdf file"""
	# Header with file path
	if color_mode: cprint("\n[+] Metadata for file: %s" % (file_full_path), "green", attrs=["bold"])
	else: print "\n[+] Metadata for file: %s" % (file_full_path)
	# Open the file
	try: 
		pdf_file = PdfFileReader(file(file_full_path, "rb"))
	except: 
		if color_mode: cprint("Could not read this file. Sorry!", "red")
		else: print "Could not read this file. Sorry!"
		return
	if pdf_file.isEncrypted: # Temporary workaround, pdf encrypted with no pass
		try: 
			pdf_file.decrypt('')
		except: 
			if color_mode: cprint("\tCould not decrypt this file. Sorry!", "red")
			else: print "\tCould not decrypt this file. Sorry!"
			return
	# Data structure with document information
	pdf_info = pdf_file.getDocumentInfo()
	# Print metadata
	if pdf_info: 
		for metaItem in pdf_info: 
			try: 
				if color_mode: 
					cprint("\t-" + metaItem[1:] + ": ", "cyan", end="")
					cprint(pdf_info[metaItem])
				else: 
					print "\t-" + metaItem[1:] + ": "
					print pdf_info[metaItem]
			except TypeError: 
				if color_mode: cprint("\t-" + metaItem[1:] + ": " + "Error - Item not readable", "red")
				else: print "\t-" + metaItem[1:] + ": " + "Error - Item not readable"
	else:
		if color_mode: cprint("\t No data found", "red")
		else: print "\t No data found"
	print ""

	
# Main function
def main (argv):
	"""Main function of the program"""
	# Check arguments
	parser = argparse.ArgumentParser(description="Metadata analyzer")
	parser.add_argument("-c", "--colored", help="activates colored output", action="store_true")
	parser.add_argument("directory", type=str, help="target directory")
	args = parser.parse_args()
	color_mode = args.colored
	target = args.directory

	printMeta(target, color_mode)
	"""
	if len(argv) != 2: 
		print "Incorrect number of arguments. "
		return
	# If arguments OK, execute function
	else: 
		target = argv[1]
		printMeta(target)
	"""


# Execute main function
if __name__ == "__main__": 
	main(sys.argv)
