import pandas as pd
from docx import Document

df = pd.read_excel("pst_art.xlsx")
dct = df.to_dict(orient='index')

j_ru = 'Петербургская социология сегодня'
j_en = 'St. Petersburg Sociology Today'
year = '2024'
issue = 25

art_nums = [v['art_num'] for k, v in dct.items() if v['issue'] == issue]
art_nums.sort()
print(art_nums)

items_lst = [v for i in art_nums for k, v in dct.items() if i == v['art_num']]


def add_content(document, content):
    section = ''
    for c in content:
        if i != c[0]:
            section = c[0]
            document.add_paragraph(section)
        p = document.add_paragraph('')
        p.add_run(c[1]).italic = True
        p.add_run(f" {c[2]}")
    document.add_page_break()


def get_fio(rec):
    iof_l = rec['iof'].split(' ')
    f_name = iof_l[0]
    surname = iof_l[1]
    s_name = iof_l[2]
    fio = f"{s_name} {f_name[0]}. {surname[0]}."
    return fio


def get_cit_ru(rec, fio):
    cit = f"{fio} {rec['title']} // {j_ru}. {year}. № {issue}. С. {rec['pp']}. "
    cit += f"DOI: {rec['doi']}; EDN: {rec['edn']}"
    return cit


def get_receiving(rec):
    receiving = f"Статья поступила в редакцию: {rec['received']}; "
    receiving += f" поступила после рецензирования и доработки: {rec['revised']}; принята к публикации: {rec['accepted']}."
    return receiving


def add_art(document, rec, cit, receiving):
    document.add_heading(rec['section'], level=2)
    document.add_paragraph(f"DOI: {rec['doi']}")
    document.add_paragraph(f"EDN: {rec['edn']}")
    document.add_paragraph(f"УДК {rec['udc']}")

    document.add_paragraph(f"{rec['iof']}1")
    document.add_paragraph(f"1 {rec['aff']}")
    document.add_heading(rec['title'], level=3)
    p = document.add_paragraph('')
    p.add_run('Аннотация.').italic = True
    p.add_run(f" {rec['abstr']}")

    document.add_paragraph(f"Ключевые слова: {rec['kw']}")
    document.add_paragraph(f"Ссылка для цитирования: : {cit}")
    document.add_paragraph(receiving)


doc = Document()
doc.add_heading('СОДЕРЖАНИЕ', level=2)

content_ru = []
for i in items_lst:
    fio = get_fio(i)
    content_ru.append((i['section'], fio, i['title']))
print(content_ru)
add_content(doc, content_ru)

for i in items_lst:
    fio = get_fio(i)
    cit = get_cit_ru(i, fio)
    receiving = get_receiving(i)
    add_art(doc, i, cit, receiving)

doc.save(f'pst_{issue}.docx')
