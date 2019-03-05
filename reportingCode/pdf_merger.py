import os
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
#
def merger(output_path, input_paths):
	pdf_merger = PdfFileMerger()
	file_handles = []
	for path in input_paths:
		pdf_merger.append(path)
	with open(output_path, 'wb') as fileobj:
		pdf_merger.write(fileobj)
#
file_list = os.listdir("C:\\Users\\2019353\\Desktop\\BlackBoxReporter\\PDF Reports")
#
