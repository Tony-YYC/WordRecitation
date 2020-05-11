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
    if meaning_list == None:
        print(word,"没有")
    for i in meaning_list:
        if not "人名" in i:
            print(word,i)
            sheet[meaningcolume+str(row)].value = i
            row+=1
            times+=1
            if times == len(meaning_list):
                row-=1
            else:
                sheet.insert_rows(row)
        else:
            sheet.delete_rows(row)
            row-=1
    return row

def indexer(she):
    rowdelete = 1
    sheet = sheetlist[she - 1]
    row = 2
    word_colume = 'A'
    meanings_colume = 'B'
    newstate = False
    while True:
        value_word = sheet[word_colume+ str(row)].value
        if newstate == True and value_word:
            row = newwordtranslater(sheet,value_word,row,meanings_colume)
        elif newstate and not value_word:
            break
        if value_word == None and newstate == False:
            pass
        elif value_word[0] ==  '#':
            if value_word == "#new":
                rowdelete = row
                newstate = True
        elif newstate == False:
            pass
        row+=1
    sheet.delete_rows(rowdelete)

i = input("请输入要翻译的表格序号:")
indexer(int(i))
forms.save("words.xlsx")
