from openpyxl import load_workbook
from translate import translate

trans = translate()

forms = load_workbook("words.xlsx")


sheetlist = []
t = 0
for sheets in forms:
    t+=1
    sheetlist.append(sheets)
    print(t,">",sheets)


def newwordtranslater(sheet,word,row,meaningcolume):
    trans.getword(word)
    meaning_list = trans.translateSep()
    times = 0
    for i in meaning_list:
        print(i)
        sheet[meaningcolume+str(row)].value = i
        print(sheet[meaningcolume+str(row)].value)
        row+=1
        times+=1
        if times == len(meaning_list):
            pass
        else:
            sheet.insert_rows(row)
    return row

def indexer(she):
    sheet = sheetlist[she - 1]
    row = 2
    word_colume = 'A'
    meanings_colume = 'B'
    newstate = False
    while row<=21:
        value_word = sheet[word_colume+ str(row)].value
        if newstate == True:
            row = newwordtranslater(sheet,value_word,row,meanings_colume)
        if value_word == None and newstate == False:
            pass
        elif value_word[0] ==  '#':
            if value_word == "#new":
                newstate = True
        elif newstate == False:
            pass
        row+=1

i = input("请输入要翻译的表格序号:")
indexer(int(i))
forms.save("words.xlsx")
