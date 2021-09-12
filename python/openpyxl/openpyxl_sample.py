import os, re
import openpyxl as px
from datetime import datetime

class ObjManager:
    # 行ごとの配列
    obj_list      = []
    # ヘッダの配列
    header_list   = []
    # ステータス集計用の配列
    cnt_predicate = [0, 0, 0]

    def __init__(self, obj):
        self.id              = obj[0].value if obj[0].value is not None else ""
        self.text            = obj[1].value if obj[1].value is not None else ""
        self.status_customer = obj[2].value if obj[2].value is not None else ""
        self.status_vendor   = obj[3].value if obj[3].value is not None else ""
        self.predicate       = self.predicate()

    def predicate(self):
        result = 0
        if self.status_customer == "TRUE" and self.status_vendor == "TRUE":
            result += 2
        elif self.status_customer == "TRUE" or self.status_vendor == "TRUE":
            result += 1
        ObjManager.cnt_predicate[result] += 1
        print(ObjManager.cnt_predicate)
        return result

    def toArray(self):
        return [self.id, self.text, self.status_customer, self.status_vendor, self.predicate]

    # クラス変数の配列を初期化する
    @classmethod
    def delArray(cls):
        ObjManager.obj_list.clear()
        ObjManager.header_list.clear()
        ObjManager.cnt_predicate.clear()
        ObjManager.cnt_predicate = cnt_predicate = [0, 0, 0]

def main():
    # Input内xlsxファイルパスをリストに格納する(Inputフォルダが必要です)
    path_list = []
    for dirpath, dirname, filenames in os.walk(os.getcwd() + "\Input"):
        for fname in filenames:
            if re.search(".xlsx", fname):
                mypath = [dirpath, fname]
                path_list.append(os.path.join(*mypath))

    # 新規でOutput.xlsxを作成する
    result_fname = "Output.xlsx"
    result_wb = px.Workbook()
    #result_wb = px.load_workbook(filename=result_fname)

    # xlsxファイルパスリストから1つずつ読み込む
    target_ws = "Sheet1"
    for fcount, fpath in enumerate(path_list, start=0):
        wb = px.load_workbook(filename=fpath)
        ws = wb[target_ws]
        print("row_num: " + str(ws.max_column))

        # 1行目をヘッダ配列に格納、それ以外は行ごとの配列を格納する
        for column_num, obj in enumerate(ws, start=1):
            if column_num == 1:
                ObjManager.header_list.append(ObjManager(obj))
            else:
                ObjManager.obj_list.append(ObjManager(obj))
        wb.save(fpath)

        # ファイルパスをシート名に設定する
        input_fname = os.path.basename(fpath)
        for s in result_wb:
            if s.title == input_fname:
                result_wb.remove(result_wb[input_fname])
        result_ws = result_wb.create_sheet(input_fname)

        # シートを読み込みオブジェクトにする
        for i, obj in enumerate(ObjManager.obj_list, start=1):
            for j, val in enumerate(obj.toArray(), start=1):
                result_ws.cell(row=i, column=j).value = val
                result_ws.cell(row=i, column=j).alignment = px.styles.Alignment(wrapText=True)

        # 出力シートの1行目にヘッダ情報を挿入する
        result_ws.insert_rows(1)
        header_zero = 0
        for j, attr in enumerate(ObjManager.header_list[header_zero].toArray(), start=1):
            result_ws.cell(row=1, column=j).value = attr
        
        # 配列を初期化する
        ObjManager.delArray()

    result_wb.save(result_fname)

if __name__ == "__main__":
    main()