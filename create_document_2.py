import string
import openpyxl
from docxtpl import DocxTemplate
import docx
import compose_doc

def get_start_row(max_row,sheet)->int:
	for _row in range(1,max_row):
		if sheet.cell(row = _row, column=1).value == '+':
			return _row

def get_start_col(max_col,sheet)->int:
	for _col in range(1,max_col):
		if sheet.cell(row = 1, column=_col).value == "Курс 1":
			return _col

def get_flag_table_semestr(wb):
	sheet = wb["План"]
	if sheet.cell(row = 2, column=1).value == None:
		return True
	else:
		return False

def counter_semestr(wb):
    sheet = wb["Титул"]
    value = sheet.cell(row = 43, column=openpyxl.utils.column_index_from_string('C')).value
    term = int(value.split(':')[1][1])
    if get_flag_table_semestr(wb):
        return term*2
    else:
        return term

def get_end_col(wb,max_col):
	sheet = wb["План"]
	if get_flag_table_semestr(wb):
		for _col in range(1,max_col):
			if sheet.cell(row = 3, column=_col).value == "Компетенции":
				return _col
	else:
		for _col in range(1,max_col):
			if sheet.cell(row = 2, column=_col).value == "Компетенции":
				return _col

def get_dict_competencies(wb):
    dict_competencies = {}
    sheet = wb["Компетенции"]
    _row = 2
    key = ''
    while (sheet.cell(row = _row, column=5).value != "exit"):
        if sheet.cell(row = _row, column=5).value != None and sheet.cell(row = _row, column=5).value != '-':
            col_test = 1
            while(sheet.cell(row = _row, column=col_test).value == None):
                col_test+=1
            key = sheet.cell(row = _row, column=col_test).value
            dict_competencies[key] = [sheet.cell(row = _row, column=4).value,{}]
        else:
            if sheet.cell(row = _row, column=5).value == '-':
                col_test = 2
                while(sheet.cell(row = _row, column=col_test).value == None):
                    col_test+=1
                dict_competencies[key][1][sheet.cell(row = _row, column=col_test).value] = sheet.cell(row = _row, column=4).value
        _row+=1
    return dict_competencies

def get_rezult_str(str_mas_cod,dict_competencies):
	mas_str = str_mas_cod.split(";")
	for index in range(len(mas_str)):
		if mas_str[index][0] == ' ':
			mas_str[index] = mas_str[index][1:len(mas_str[index])]

	str1 = ''
	str2 = ''
	dict_competencies_final = {}
	for key in dict_competencies.keys():
		for key2 in dict_competencies[key][1].keys():
			dict_competencies_final[key2] = [dict_competencies[key][1][key2],key+' '+dict_competencies[key][0]]
	for key in mas_str:
		str2+= key+' '+ dict_competencies_final[key][0]+'\n\t'
		if not(dict_competencies_final[key][1] in str1):
			str1 += dict_competencies_final[key][1] + '\n\t'

	str1 = str1[:len(str1)-2]
	str2 = str2[:len(str2)-2]

	return str1,str2

def get_form_control(sheet,_row):
	final_str = ''
	mas_form_control = ['Экзамен','Зачет','Зачет с оценкой']
	mas = []
	for index in range(4,7):
		all_semester = sheet.cell(row = _row, column=index).value
		if all_semester != None:
			for semester in all_semester:
				mas.append("Семестр " + semester+ ', ' + mas_form_control[index-4])
	
	for str_f_m in sorted(mas,key = lambda s: int(s[8])):
		final_str+=str_f_m+'\n\t'

	return final_str[:len(final_str)-2]

def get_sum_L(sheet,_row,_col,counter_semestr):
	sum_L = 0
	_col = _col +1
	for _ in range(counter_semestr):
		value = sheet.cell(row = _row, column=_col).value
		if value != None:
			sum_L+=float(value)
		_col+=10
	return str(sum_L)

def get_sum_C(sheet,_row,_col,counter_semestr):
	sum_C = 0
	_col = _col + 4
	for _ in range(counter_semestr):
		value = sheet.cell(row = _row, column=_col).value
		if value != None:
			sum_C+=float(value)
		_col+=10
	return str(sum_C)

def get_sum_LR(sheet,_row,_col,counter_semestr):
	sum_LB = 0
	_col = _col + 2
	for _ in range(counter_semestr):
		value = sheet.cell(row = _row, column=_col).value
		if value != None:
			sum_LB+=float(value)
		_col+=10
	return str(sum_LB)

def get_sum_P(sheet,_row,_col,counter_semestr):
	sum_P = 0
	_col = _col + 3
	for _ in range(counter_semestr):
		value = sheet.cell(row = _row, column=_col).value
		if value != None:
			sum_P+=float(value)
		_col+=10
	return str(sum_P)

def read_word_fila(name_fila,name_d):
	doc = docx.Document(name_fila)
	all_paras = doc.paragraphs
	index_start = -1
	text = ""
	for text_index in range(len(all_paras)):
		if name_d in all_paras[text_index].text:
			index_start = text_index
			break

	if index_start != -1:
		for text_index_2 in range(index_start,len(all_paras)):
			if all_paras[text_index_2].text == "Тематический план:":
				index_start = text_index_2+1
				break

		for text_index_3 in range(index_start,len(all_paras)):
			if not("Тема" in all_paras[text_index_3].text):
				break
			text+='\t'+all_paras[text_index_3].text+'\n'

	if text == "":
		text = "Здесь должны быть темы"
	return text

def get_dict_key_all_docs(name_table):
	final_dict = {}
	wb = openpyxl.load_workbook(name_table)
	dict_competencies = get_dict_competencies(wb)
	sheet = wb["Титул"]
	final_dict["SP"]=sheet.cell(row = 29, column=openpyxl.utils.column_index_from_string('D')).value
	final_dict["PP"]=sheet.cell(row = 30, column=openpyxl.utils.column_index_from_string('D')).value
	final_dict["FO"]= (sheet.cell(row = 42, column=openpyxl.utils.column_index_from_string('C')).value).split(' ')[2]
	final_dict["KV"] = (sheet.cell(row = 40, column=openpyxl.utils.column_index_from_string('C')).value).split(' ')[1]
	final_dict["G"] = sheet.cell(row = 40, column=openpyxl.utils.column_index_from_string('W')).value

	sheet = wb["План"]
	max_rows = sheet.max_row+1
	max_cols = sheet.max_column+1
	col_start = get_start_col(max_cols,sheet)
	count_semestr = counter_semestr(wb)
	end_col = get_end_col(wb,max_cols)
	
	for _row in range(get_start_row(max_rows,sheet),max_rows):
		if sheet.cell(row = _row, column=end_col-1).value != None:
			final_dict["NAME"]= sheet.cell(row = _row, column=3).value
			print(final_dict["NAME"])
			final_dict["CODE"] = sheet.cell(row = _row, column=2).value
			final_dict["DEVELOP_GOAL"],final_dict["DEVELOP_REZULT"] = get_rezult_str(sheet.cell(row = _row, column=end_col).value,dict_competencies)
			mas_places_dissiple = ['Дисциплина относится к обязательной части образовательной программы.','Дисциплина относится к части образовательной программы, формируемой участниками образовательных отношений, является обязательной для изучения.','Дисциплина относится к части образовательной программы, формируемой участниками образовательных отношений, предлагается обучающимся на выбор.']
			if 'О' in final_dict["CODE"]:
				final_dict["PLACE"] = mas_places_dissiple[0]
			elif ('ДВ' or 'ФТД') in final_dict["CODE"]:
				final_dict["PLACE"] = mas_places_dissiple[2]
			else:
				final_dict["PLACE"] = mas_places_dissiple[1]

			final_dict["LIGHTING_SEMES"] = get_form_control(sheet,_row)
			final_dict["ZE"] = sheet.cell(row = _row, column=openpyxl.utils.column_index_from_string('H')).value
			final_dict["K"] = sheet.cell(row = _row, column=openpyxl.utils.column_index_from_string('K')).value
			final_dict["L"] = get_sum_L(sheet,_row,col_start,count_semestr)
			final_dict["C"] = get_sum_C(sheet,_row,col_start,count_semestr)
			final_dict["P"] = get_sum_P(sheet,_row,col_start,count_semestr)
			final_dict["LR"] = get_sum_LR(sheet,_row,col_start,count_semestr)
			mas_path = [".\\Resources\\Шаблоны\\Общий черновик аннотаций иностранцы.docx",".\\Resources\\Шаблоны\\Общий черновик аннотаций.docx"]
			if len(final_dict["NAME"].split("*")) > 1:
				final_dict["CONTENT"] = read_word_fila(mas_path[0],final_dict["NAME"].split("*")[0][:-1])
			else:
				final_dict["CONTENT"] = read_word_fila(mas_path[1],final_dict["NAME"])
			yield final_dict
		else:
			continue


def main():
	name_table = input("Введите полный путь к плану")
	key_plan = ["NAME","SP","PP","FO","KV","G","CODE","DEVELOP_GOAL","DEVELOP_REZULT","PLACE","LIGHTING_SEMES","ZE","K","L","C","P","LR","CONTENT"]
	key_FOS = ["NAME","SP","PP","G","DEVELOP_GOAL","DEVELOP_REZULT"]
	for document in get_dict_key_all_docs(name_table):
		docPlan = DocxTemplate(".\\Resources\\Шаблоны\\Template.docx")
		contex_plan = {}
		for key in key_plan:
			contex_plan[key] = document[key]
		docPlan.render(contex_plan)
		if "*" in document["NAME"]:
			docPlan.save("doc/plan/"+document["NAME"].replace('/','_').replace('*','_')+".docx")
		else:
			docPlan.save("doc/plan/"+document["NAME"].replace('/','_')+".docx")

		docFOS = DocxTemplate(".\\Resources\\Шаблоны\\2022_Шаблон ФОСа.docx")
		contex_FOS ={}
		for key in key_FOS:
			contex_FOS[key] = document[key]
		
		docFOS.render(contex_FOS)
		if "*" in document["NAME"]:
			docPlan.save("doc/FOS/"+document["NAME"].replace('/','_').replace('*','_')+".docx")
		else:
			docPlan.save("doc/FOS/"+document["NAME"].replace('/','_')+".docx")

	compose_doc.combine_all_docx("asd.docx",name_table.split('\\')[1].split(".plx")[0])
if __name__ == '__main__':
	main()






















