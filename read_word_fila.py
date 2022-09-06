import docx



doc = docx.Document("Resources/Общий черновик аннотаций.docx")

all_paras = doc.paragraphs


print(len(all_paras))
#for para in all_paras:
#print(all_paras[0].text)

second_para = all_paras[9]
for run in second_para.runs:
    print(run.text)