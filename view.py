# _*_ coding:utf-8 _*_
import random
# from translate import translate
from openpyxl import load_workbook

class WordModel:
    """
    word指单词本身
    meanings 指单词释义
    score指单词得分情况
    """
    def __init__(self,row = 0,word = '' ,sheet = None , meaning = '',father = None,listid = -1):
        self.row = row
        self.word = word
        self.meaning = meaning
        self.sheet = sheet
        self.father = father
        self.listid = listid
class Controller:
    """

    """

    rowlist = []
    sheetlist = []

    def __init__(self,username = ''):
        self.username = username
        self.wordlist = []

    def input_name(self,name):
        self.username = name
    def output_name(self):
        return self.username

    def add_word(self,word):
        self.wordlist.append(word)

    @property
    def get_word_list(self):
        return self.wordlist

    def get_all_meaning(self,i):
        father = i.father
        if father:
            for wordrow in range(father.listid,self.wordlist[-1].listid):
                if self.wordlist[wordrow].word == father.word:
                    print(self.wordlist[wordrow].meaning)
                else:
                    break
        else:print('这个单词没有提示。。。')


    def randomize_testsheet(self,sheet):
        list_for_random = []

        if sheet != 0:
            flag = 0
            for i in self.wordlist:
                if i.sheet == sheet:
                    list_for_random.append(i)
                    flag = 1
                elif i.sheet != sheet and flag == 1:
                    break
                elif i.sheet != sheet and flag == 0:
                    pass
        else:
            for i in self.wordlist:
                list_for_random.append(i)
        random.shuffle(list_for_random)
        return list_for_random

    def choice_appender(self,word,testlist,choices = 4):
        choice_list = []
        choice_list.append(word)
        while len(choice_list) < choices:
            i = testlist[random.randint(1,len(testlist))-1]
            if i.word != word.word and not i in choice_list:
                choice_list.append(i)
            else:
                pass
        random.shuffle(choice_list)
        return choice_list


class View:
    def __init__(self):
        self.__controller = Controller()

    def main(self):
        self.importer()
        while True:
            self.print_menu()
            self.__select_menu_item()
            # os.system("cls")

    def importer(self):
        for sheet in sheetlist:
            row = 2
            word_colume = element_locater('#带井号的不会统计', sheet)
            meanings_colume = element_locater('#解释', sheet)
            while blank_checker(sheet,row):
                value_word = sheet[word_colume+ str(row)].value

                # print(value_word)
                #测试

                if value_word == None:
                    word = WordModel()
                    word.word = self.__controller.get_word_list[-1].word
                    word.row = row
                    word.meaning = sheet[meanings_colume + str(row)].value
                    word.sheet = sheet
                    word.father = self.__controller.get_word_list[-1]
                    word.listid = len(self.__controller.get_word_list)
                    self.__controller.add_word(word)
                elif value_word[0] ==  '#':
                    pass
                else:
                    word= WordModel()
                    word.word = value_word
                    word.row = row
                    word.meaning = sheet[meanings_colume + str(row)].value
                    word.sheet = sheet
                    word.listid = len(self.__controller.get_word_list)
                    self.__controller.add_word(word)

                row += 1

    def print_menu(self):
        print("1)登陆或注册")
        print("2)测试")
        print("3)学习")
        print("4)显示所有单词")
        print("5)退出")
        # print(self.__controller.wordlist[2].word)

    def __select_menu_item(self):
        item = input("请您输入选项:")
        if item == "1":
            self.login()
        elif item == "2":
            self.test()
        elif item == "3":
            self.learn()
        elif item == "4":
            self.show_word()
        elif item == "5":
            forms.save("单词本.xlsx")
            forms.close()
            exit(0)
        else:
            print("输入错误")


    def login(self):
        name = input("输入用户名")
        self.__controller.input_name(name)

    @property
    def select_sheet(self):
        for mn in range(len(sheetlist)):
            print(sheetlist[mn])
        ss = input("选择第几个表格或者输入a全选")
        if ss == 'a':
            return 0
        else:
            return sheetlist[int(ss)-1]

    @property
    def checker(self):
        if self.__controller.output_name() == '':
            print("请先登陆")
            return False
        else:
            return True

    def test(self):
        print('============================')
        if self.checker:
            i = self.select_sheet
            testlist = self.__controller.randomize_testsheet(i)
            if input("输入1来根据单词选词义，2来根据词义默写单词") == '1':
                self.test_by_choose_meaning(testlist)
            else:self.test_by_enter_word(testlist)

    def test_by_choose_meaning(self,testlist,repeat = True):
        print("注意一定要大写字母选项，只能填 ABCD等等   输入0来退出")
        for i in testlist:
            print(i.word)
            chioce_list = self.__controller.choice_appender(i, testlist, 4)
            for m in range(len(chioce_list)):
                print("%s:%s" % (chr(65 + m), chioce_list[m].meaning))
            chioce = input("请选择：")
            if chioce == '0':
                break
            chioce = ord(chioce) - 65
            if chioce_list[chioce].word == i.word or chioce_list[chioce].meaning == i.meaning:
                print('=========================')
                print("恭喜你答对了")
                print('word:      ', i.word)
                print('meanings:  ', i.meaning)
                print("=========================")
            else:
                print("=========================")
                print("答错了。。。正确答案为：%s" % (i.meaning))
                print("=========================")
                if repeat : testlist.append(i)

    def test_by_enter_word(self,testlist,repeat = True):
        for i in testlist:
            print(i.meaning)
            answer = input("请输入单词，输入1可以获取提示，输入0退出")
            if answer == i.word:
                print('=========================')
                print("恭喜你答对了")
                print('word:      ', i.word)
                print('meanings:  ', i.meaning)
                print("=========================")
            elif answer == '0':
                break
            elif answer =='1':
                self.__controller.get_all_meaning(i)
                answer = input("现在呢？会了吗？手动滑稽")
                if answer == i.word:
                    print('=========================')
                    print("恭喜你答对了")
                    print('word:      ', i.word)
                    print('meanings:  ', i.meaning)
                    print("=========================")
                else:
                    print("=========================")
                    print("答错了。。。正确答案为：%s" % (i.word))
                    print("=========================")
                    if repeat: testlist.append(i)
            else:
                print("=========================")
                print("答错了。。。正确答案为：%s" % (i.word))
                print("=========================")
                if repeat : testlist.append(i)


    def learn(self):
        if self.checker:
            i = self.select_sheet
            if i == 0:
                pass
            else:
                pass

    def show_word(self):
        wordlist = self.__controller.get_word_list
        for i in wordlist:
            print('word:      ' ,  i.word)
            print('meanings:  ' , i.meaning)
            print('sheet:  ' ,  i.sheet)
            print('row:       ' ,  i.row)
            print("=========================")

userlist = []
sheetlist = []
print("initializing........")
# localpath = os.getcwd()
# dirinformation = os.listdir(localpath)
# print(dirinformation)
forms = load_workbook("单词本.xlsx")
print("找到单词本：")

def initialize_user(sheet,listnumber):
    for ascid in range(ord('A'),ord('Z')):
        cellname = chr(ascid)+'1'
        # print(cellname)
        # print(sheet[cellname].value)
        if sheet[cellname].value == '#单个人':
            userlist[listnumber][sheets].append(ascid)
            for b in range(ascid,ord('Z')):
                cellname_2 = chr(b)+'2'
                if sheet[cellname_2].value != None:
                    userlist[listnumber][sheets].append(sheet[cellname_2].value)
                else:
                    break
            break

def element_locater(element,sheet):
    for ascid in range(ord('A'),ord('Z')):
        cellname = chr(ascid)+'1'
        # print(cellname)
        # print(sheet[cellname].value)
        if sheet[cellname].value == element:
            return chr(ascid)

def blank_checker(sheet,row):
    bo = True
    for ascid in range(ord('A'), ord('F')):
        cellname = chr(ascid) + str(row)
        if sheet[cellname].value != None:
            return True
        else:
            bo = False
    return bo
a = 0
for sheets in forms:
    sheetlist.append(sheets)
    userlist.append({})
    userlist[a][sheets]=[]
    print(sheets)
    initialize_user(sheets,a)
    # print(userlist)
    a+=1


# well = translate()
# well.getword("well")
# print(well.translation())
#测试translate
View().main()
