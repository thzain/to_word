import xlrd
from pandas import DataFrame


class ColValue:
    def __init__(self):
        self.seq = ""
        self.item = ""
        self.unit = ""
        self.requirement = ""
        self.result = ""
        self.assessment = ""

class sheetData:
    def __init__(self):
        self.head = []
        self.header_data = []
        self.body_data = []
        self.df = DataFrame()


class Xlsx:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheets = []
        self.sheets_data = []
        self.read_xlsx()

    def read_xlsx(self):
        # 获取sheets名
        workbook = xlrd.open_workbook(self.file_path)
        self.sheets = workbook.sheet_names()
        # 获取sheets数据
        for sheet in self.sheets:
            sheet_data = workbook.sheet_by_name(sheet)
            # 获取行数与列数
            nrows = sheet_data.nrows
            ncols = sheet_data.ncols
            # 转为列表
            sheet_all = sheetData()
            body_rows = []
            for row in range(nrows):
                row_data = sheet_data.row_values(row)

                # 都转为字符串
                row_data = [str(i) for i in row_data]

                # 如果 row_data 的长度小于 ncols，补全
                if len(row_data) < ncols:
                    row_data.extend([''] * (ncols - len(row_data)))
                if row == 0:
                    sheet_all.head = row_data
                if row == 1:
                    sheet_all.head_data = row_data
                if row > 1:
                    body_rows.append(row_data)
            sheet_all.body_data = body_rows
            sheet_all.df = DataFrame(sheet_all.body_data, columns=sheet_all.head)
            self.sheets_data.append({sheet: sheet_all})


def clean_data(file_path_in):
    excel_data = Xlsx(file_path_in)

    # 用一个二维列表存储最终需要输出的数据
    res_data = []

    # 先写入两个空行
    col_empty = ColValue()
    res_data.append(col_empty)
    res_data.append(col_empty)

    # 写入 2 传输特性的空行
    col_2 = ColValue()
    col_2.seq = '2'
    col_2.item = '传输特性'
    res_data.append(col_2)

    # 下面写入其他数据
    seq_new = 2.1
    for sheet in excel_data.sheets_data[2:]:
        # 如果是新的表格，需要重新定义 seq
        seq_new_case = True
        for key, value in sheet.items():
            # 获取项目
            item = value.head[0]
            # 获取值列的单位
            col_unit = value.head_data[1].split('[')[1]
            col_unit = col_unit.split(']')[0]

            # 获取单位列
            col_requirement = value.head[1]

            # 获取结果列
            body_data = value.body_data
            for index, row in enumerate(body_data):
                if seq_new_case:
                    col = ColValue()
                    # 保留一个小数
                    seq_new_str = round(seq_new, 1)
                    col.item = item
                    col.seq = str(seq_new_str)
                    seq_new += 0.1
                    seq_new_case = False
                    res_data.append(col)

                col = ColValue()

                col.item = row[1] + ' ' + col_unit
                col.unit = col_requirement
                col.requirement = ""
                col.result = row[3]
                col.assessment = 'P'
                res_data.append(col)

    # 一个sheet做完，空2行
    res_data.append(col_empty)
    res_data.append(col_empty)

    return res_data


if __name__ == "__main__":
    file_path = './data/20240905-原始测试数据.xls'
    res = clean_data(file_path)

