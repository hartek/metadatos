#!/usr/bin/python
# -*- encoding: utf-8 -*- 

from PyPDF2 import PdfFileReader,PdfFileWriter # Module for PDF manipulation
import docx # Module for DOCX manipulation
import os # Module for OS interaction
import sys # Module for system interaction
from termcolor import cprint # Module for terminal colored output
import datetime

"""
CREATED BY GUILLERMO ROMAN FERRERO a.k.a. hartek
This script prints the metadata of the PDF files found recursively in a directory given by argument
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
			else: 
				pass

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
