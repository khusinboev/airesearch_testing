from docx import Document
from docx.enum.style import WD_STYLE_TYPE


def clean_text(text):
    cleaned_text = ' '.join(text.split())
    return cleaned_text


# 33
file = "L 33 Англо-русско-узбекский краткий словарь тер-минов по распространению радиоволн и антенно-фидерным устройствам. 2-часть.docx"
doc = Document(docx=file)

r, p, q, w = 0, 0, 0, 0
tables = doc.tables
for i in tables:
    for row in i.rows:
        tex = str(row.cells[0].text)
        desc = str(row.cells[-1].text)
        if row.cells[-1].text == row.cells[0].text or (row.cells[0] == '' and row.cells[-1] == ''):
            pass

        elif "en -" in row.cells[0].text and "uz -" in row.cells[0].text:
            term_uz = tex[tex.index('uz -'):tex.index('en -')][5:].strip()
            term_en = tex[tex.index('en -'):][5:].strip()
            term_ru = tex[:tex.index('uz -')].strip()
            descs = [element for element in desc.split('\n\n') if element]
            if len(descs) == 3:
                desc_uz = desc.split('\n\n')[1]
                desc_uz_kr = desc.split('\n\n')[2]
                desc_ru = desc.split('\n\n')[0]

                if len(term_uz.split('\n ')) == 2:
                    term_uz_ln = term_uz.split('\n ')[0]
                    term_uz_kr = term_uz.split('\n ')[1]
                    r += 1
                else:
                    w += 1
                    print(desc)
                q += 1
            else:
                pass
        else:
            pass
print(q)
print(r)
print(p)
print(w)
