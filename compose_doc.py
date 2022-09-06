from docxcompose.composer import Composer
from docx import Document as Document_compose
import os

def combine_all_docx(filename_master,fila_name):
	directory_list = list()

	for root, dirs, files in os.walk("doc", topdown=False):
		for name in files:
			directory_list.append(os.path.join(root, name))

	print(directory_list)
	number_of_sections=len(directory_list)
	master = Document_compose(filename_master)
	composer = Composer(master)
	for i in range(0, number_of_sections):
		doc_temp = Document_compose(directory_list[i])
		composer.append(doc_temp)
	composer.save(fila_name+".docx")