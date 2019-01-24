from docxtpl import DocxTemplate
from shutil import copyfile
import sqlite3

conn = sqlite3.connect("bd.db")
curs = conn.cursor()

curs.execute("""CREATE TABLE organisations
                  (name_comp text, ogrn integer, inn integer, kpp integer, ur_address text, fakt_address text,
                   post_address text, rs integer, ks integer, bik integer, oktmo integer, okved2 integer)
               """)
conn.commit()

organisations = [('Bob', '11111111111', '22222222222', '333333333333', 'Green Street, 555', 'Yellow Street, 33', 'Yellow Street, 33', '4444444444', '555555555', '66666', '123', '456'),
                 ('Lucky', '123456789', '987654321', '755238', 'Sunshine Street, 777', 'Kitty Street, 64', 'Kitty Street, 64', '7635283464', '874523786452', '873675', '377', '736')]

curs.executemany("INSERT INTO organisations VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", organisations)
conn.commit()

sql = "SELECT * FROM organisations WHERE name_comp=?"
curs.execute(sql, [("Lucky")])
data = curs.fetchall()[0]

row = list(map(lambda x: x[0], curs.description))

copyfile('doc.docx', 'doc1.docx')
doc = DocxTemplate("doc1.docx")

context = {row[i]: data[i] for i in range(len(row))}
doc.render(context)
doc.save("doc1.docx")