import urllib.request
import json
import urllib.parse
import time
class traslate:
    def __init__(self,word):
        self.word = word
    def fetch(self,query_str):
        query = {'q': "".join(query_str)}   # list --> str: "".join(list)
        url = 'https://fanyi.youdao.com/openapi.do?keyfrom=11pegasus11&key=273646050&type=data&doctype=json&version=1.1&' + urllib.parse.urlencode(query)
        response = urllib.request.urlopen(url, timeout=3)
        html = response.read().decode('utf-8')
        return html
    
    def parse(self,html):
        d = json.loads(html)
        try:
            if d.get('errorCode') == 0:
                explains = d.get('basic').get('explains')
                result = str(explains).replace('\'', "").replace('[', "").replace(']', "")  #.replace真好用~
                return result
            else:
                print('无法翻译!****')
                return ""       #若无法翻译，则空出来
        except:
            return ""      #若无法翻译，则空出来
    def translation(self):
        chinese = self.parse(self.fetch(self.word))
        return chinese
    def translateSep(self):
        string = self.translation().split(",")
        return string

word = input("测试状态，输入单词翻译")
tran = traslate(word)
print(tran.translation())
print(tran.translateSep())
#translation returns a word while translateSep returns a list