from openpyxl import load_workbook
from docxtpl import DocxTemplate



def convert():
    doc = DocxTemplate('tabl.docx')

    book = load_workbook(filename= 'test2.xlsx')
    sheet = book['WeekJournal']

    clas = {}

    for name in range(1, 35):
        clas[sheet[f'b{name+10}'].value] = {}
        for day in 'cdefghijklmnopqrstuvw':
            if sheet[f'{day}6'].value != None:
                day_obozn = None
                if 'понедельник' in sheet[f'{day}6'].value:
                    day_obozn = 'пн'
                elif 'вторник' in sheet[f'{day}6'].value:
                    day_obozn = 'вт'
                elif 'среда' in sheet[f'{day}6'].value:
                    day_obozn = 'ср'
                elif 'четверг' in sheet[f'{day}6'].value:
                    day_obozn = 'чт'
                elif 'пятница' in sheet[f'{day}6'].value:
                    day_obozn = 'пт'
                clas[sheet[f'b{name+10}'].value][f'{day_obozn}'] = sheet[f'{day}{name+10}'].value

    contents = {}

    for i, name in zip(range(1, len(clas)+1), clas):
        if name == None:
            contents[f'fio{i}'] = ''
        else:
            contents[f'fio{i}'] = name

    for day in ['пн', 'вт', 'ср', 'чт', 'пт']:
        contents1 = {}
        for i, name in zip(range(1, len(clas)+1), clas):
            if clas[name][day] == None:
                contents1[f'{day}{i}'] = ''
            else:
                contents1[f'{day}{i}'] = clas[name][day]
        contents.update(contents1)


    doc.render(contents)
    doc.save('tabl_final.docx')






def main():
    convert()

if __name__ == "__main__":
    main()