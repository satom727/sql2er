import os
import xlsxwriter

# ER図を作るスクリプトfor MySQL
# xlsxwriter リファレンス
# 新しいファイルとワークシートを作成
#workbook = xlsxwriter.Workbook('demo.xlsx')
#worksheet = workbook.add_worksheet()
# 列Aの幅を変更
# worksheet.set_column('A:A', 20)
# 太字にする書式を追加
# bold = workbook.add_format({'bold': True})
# テキストの書き込み
# worksheet.write('A1', 'Hello')
# テキストの書き込み・書式の適用
# worksheet.write('A2', 'World', bold)
# 数値の書き込み（セル番地を数字で指定）
# worksheet.write(rowPos, columnPos, 文字)
# worksheet.write(2, 0, 123)
# worksheet.write(3, 0, 123.456)
# 画像を挿入
# worksheet.insert_image('B5', 'logo.png')

# ------class------
class Column():
    """カラムクラス comment構文非対応"""
    def __init__(self, parentTable):
        self.parentTable = parentTable
        self.name = ''
        self.dataType = ''
        self.maxLen = ''
        self.default = ''
        self.nullAbleFlg = True
        self.pkFlg = False
        self.preWord = ''
        self.sortNo = -1
    def setPK(self):
        self.pkFlg = True
    def setDef(self,param,columnCnt):
        """カラム定義設定"""
        if param.strip() == '(' or param.strip() ==')':
            pass
        elif self.name == '':
            self.name = param
            self.sortNo = columnCnt
            columnCnt += 1
        elif self.dataType == '':
            bracketBgnPos = param.find('(')
            bracketEndPos = param.rfind(')')
            if bracketBgnPos >= 0 and bracketEndPos >= 0:#()あったらmaxLen指定
                self.maxLen = param[bracketBgnPos + 1:param.rfind(')')-bracketBgnPos-1]
                self.dataType = param[0:bracketBgnPos]
            else :
                self.dataType = param
        elif self.preWord.lower() == 'default':
            self.default = param
        elif self.preWord.lower() == 'not' and param.lower() == 'null':
            self.nullAbleFlg = False
        else :
            pass
        self.preWord = param
        return columnCnt

class Table():
    """テーブルクラス"""
    def __init__(self,name=''):
        """初期化(第1引数self必須)"""
        self.name = name
        self.columnList = {}
        self.pkList = []
        self.columnCnt = 0
    def setTableName(self,name):
        self.name = name
    def addColumn(self,column):
        self.columnList[column.name] = column
    def addPK(self,column):
        if column in self.columnList:
            self.pkList.append(column)
            self.columnList[column].setPK()
    def query2Table(self,query):
        """クエリを解析しテーブル定義を作成"""
        if query.lower().find('create table') >= 0:
            wordList = query.replace(',',' , ').replace('(',' ( ').replace(')',' ) ').split()
            preWord = ''
            columnFlg = False
            pkFlg = False
            for word in wordList:
                if word.lower() == 'primary':
                    columnFlg = False
                    pkFlg = True
                elif pkFlg and not (preWord.lower() == 'primary' and word.lower() == 'key'):
                    self.addPK(word)
                    if word.strip() == ')':
                        pkFlg = False
#                    else:#デバッグ用
#                        if self.name == '`m_info`':
#                            print(column.sortNo)
#                            print(word)
                elif word == ',':
                    self.addColumn(column)#PKがなかった場合はcolumnが追加されない
                    column = Column(self.name)#Primary key のあとにユニークキーとかあると余計なカラム変数も追加される
                elif columnFlg:
                    self.columnCnt = column.setDef(word,self.columnCnt)
                elif preWord.lower() == 'table':
                    columnFlg = True
                    self.setTableName(word)
                    column = Column(self.name)
                else :
                    pass
                preWord = word

class Schema():
    """スキーマクラス"""
    def __init__(self,schemaName):
        self.schemaName = schemaName
        self.tableList = {}
    def addTable(self,table):
        self.tableList[table.name] = table


# ------関数------
def makeQueryList(file):
    """クエリごとに格納されたリストを作成する関数"""
    txt = ''
    lineList = []
    for line in file:
        line = line.lstrip()
        # コメントアウトを除く全行
        if not line.startswith('--') and not line.startswith('/*'):
            txt = txt + line
        # ;があった場合はlistに追加してtxtを初期化
        if txt.rfind(';') >= 0:
            lineList.append(txt.replace('\n',' ').lstrip())
            txt = ''
    return lineList

def makeExcelFile(schema):
    """excel出力関数"""
    tableNamePos = 0
    rowAddVal = tableNamePos + 1
    maxRow = 0
    columnPos = 0
    maxColomunNo = 26

    book = xlsxwriter.Workbook('ER図.xlsx')
    sheet1 = book.add_worksheet('ER')
    sheet1.set_column('A:Z', 20)
    titFmt = book.add_format({'bg_color': '#E0FFFF','border':1})
    pkFmt = book.add_format({'bg_color': '#FFFF00','border':1})
    fmt = book.add_format({'border':1})
    for tableObj in schema.tableList.values():
        #rowPos = 0
        if maxColomunNo <= columnPos:
            columnPos = 0
            tableNamePos = maxRow + 2
            rowAddVal = tableNamePos + 1
        sheet1.write(tableNamePos, columnPos, tableObj.name,titFmt)
        for columnObj in tableObj.columnList.values():
            if maxRow < (columnObj.sortNo + rowAddVal):
                maxRow = columnObj.sortNo + rowAddVal
            if columnObj.pkFlg:
                if columnObj.sortNo >= 0:#余計な値予防
                    sheet1.write(columnObj.sortNo + rowAddVal, columnPos, columnObj.name,pkFmt)
            else :
                if columnObj.sortNo >= 0:
                    sheet1.write(columnObj.sortNo + rowAddVal, columnPos, columnObj.name,fmt)
            #rowPos += 1
        columnPos += 2
    book.close()

# ------main処理------

# scriptDir = os.path.dirname(__file__)
inputFile = os.path.join(os.path.dirname(__file__), 'mysql_dump.sql')
with open(inputFile,'r',encoding='utf-8') as f:
    queryList = makeQueryList(f)

schema = Schema('test_schema')
for query in queryList:
    table = Table()
    table.query2Table(query)
    if(table.name != ''):
        schema.addTable(table)

makeExcelFile(schema)