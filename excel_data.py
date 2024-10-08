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

    # 加一个空行
    res_data.append(col_empty)

    # 下面写入其他数据
    seq_new = 2.1
    for sheet in excel_data.sheets_data[2:]:

        # 情况1  工作表有几个数据，就打印几行；（衰减和相时延）
        if "衰减" in sheet:
            # 如果是新的表格，需要重新定义 seq
            seq_new_case = True
            for key, value in sheet.items():
                # 获取项目
                item = value.head[0]
                # 获取值列的单位
                col_unit = value.head_data[1].split('[')[1]
                col_unit = col_unit.split(']')[0]

                # 获取单位列
                col_unit_single = value.head[1]
                col_req = value.head_data[2]

                col_fill_str = "最大"
                if "下限" in col_req:
                    col_fill_str = "最小"

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
                    col.unit = col_unit_single
                    col.requirement = (col_fill_str if "最" not in row[2] else "") + row[2]
                    col.result = row[3]
                    col.assessment = 'P'
                    res_data.append(col)

            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

        # 情况1  工作表有几个数据，就打印几行；（衰减和相时延）
        elif "相时延" in sheet:
            # 如果是新的表格，需要重新定义 seq
            seq_new_case = True
            for key, value in sheet.items():
                # 获取项目
                item = value.head[0]
                # 获取值列的单位
                col_unit = value.head_data[1].split('[')[1]
                col_unit = col_unit.split(']')[0]

                # 获取单位列
                col_unit_single = value.head[1]
                col_req = value.head_data[2]

                col_fill_str = "最大"
                if "下限" in col_req:
                    col_fill_str = "最小"

                # 获取结果列
                body_data = value.body_data
                for index, row in enumerate(body_data):
                    if seq_new_case:
                        col = ColValue()
                        # 保留一个小数
                        seq_new += 0.1
                        seq_new_str = round(seq_new, 1)
                        col.item = item
                        col.seq = str(seq_new_str)
                        seq_new_case = False
                        res_data.append(col)

                    col = ColValue()

                    col.item = row[1] + ' ' + col_unit
                    col.unit = col_unit_single
                    col.requirement = (col_fill_str if "最" not in row[2] else "") + row[2]
                    col.result = row[3]
                    col.assessment = 'P'
                    res_data.append(col)

            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

        # 情况3 如果读到的工作表的名称==‘时延差’或者工作表的名称 == ‘特性阻抗’，则去读取最差情形汇总表里的‘时延差’数据或特性阻抗的最差情形，并打印。
        elif "时延差" in sheet:
        # if True:
            worst_sheet = excel_data.sheets_data[1]
            row_data = worst_sheet['最差情形汇总'].body_data

            # 筛选第一列包含 时延差 的行
            row_data = [row for row in row_data if '时延差' in row[0]]

            col_worst = ColValue()
            seq_new = seq_new + 0.1
            col_worst.seq = round(seq_new, 1)
            col_worst.item = '时延差'
            col_worst.requirement = "最大{}".format(round(float(row_data[0][7]), 1))
            col_worst.result = round(float(row_data[0][5]), 1)
            col_worst.unit = row_data[0][0].split('[')[1].split(']')[0]
            col_worst.assessment = 'P'
            res_data.append(col_worst)

            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

        elif "特性阻抗" in sheet:
        # if True:
            worst_sheet = excel_data.sheets_data[1]
            row_data = worst_sheet['最差情形汇总'].body_data

            # 筛选第一列包含 时延差 的行
            row_data = [row for row in row_data if '特性阻抗' in row[0]]

            # 在数据表查找频率范围
            detail_data = sheet['特性阻抗']
            detail_data_body = detail_data.body_data
            detail_data_body_col_2 = [float(row[1]) for row in detail_data_body]
            detail_data_head = detail_data.head_data

            range_str = ""
            # 取出第二列的最大值和最小值
            if len(detail_data_body) > 0:
                min_2 = round(min(detail_data_body_col_2), 0)
                max_2 = round(max(detail_data_body_col_2), 0)
                unit_2 = detail_data_head[1].split('[')[1].split(']')[0]
                range_str = "({}-{}{})".format(min_2, max_2, unit_2)

            unit_mhz = "MHz"
            col_worst = ColValue()
            seq_new += 0.1
            col_worst.seq = round(seq_new, 1)
            col_worst.item = '特性阻抗{}'.format(range_str)

            res_data.append(col_worst)
            # 第二行
            col_second = ColValue()
            col_second.seq = ""
            col_second.item = "上限最差值{}{}".format(row_data[0][2], unit_mhz)
            col_second.requirement = "最大{}".format(round(float(row_data[0][7]), 0))
            col_second.result = round(float(row_data[0][5]), 1)
            col_second.unit = row_data[0][0].split('[')[1].split(']')[0]
            col_second.assessment = 'P'
            res_data.append(col_second)

            # 第三行
            col_third = ColValue()
            col_third.seq = ""
            col_third.item = "下限最差值{}{}".format(row_data[0][2], unit_mhz)
            col_third.requirement = "最小{}".format(round(float(row_data[0][3]), 0))
            col_third.result = round(float(row_data[0][1]), 1)
            col_third.unit = row_data[0][0].split('[')[1].split(']')[0]
            col_third.assessment = 'P'
            res_data.append(col_third)

            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

        elif "输入阻抗" in sheet:
        # if True:
            worst_sheet = excel_data.sheets_data[1]
            row_data = worst_sheet['最差情形汇总'].body_data

            # 筛选第一列包含 时延差 的行
            row_data = [row for row in row_data if '输入阻抗' in row[0]]

            # 在数据表查找频率范围
            detail_data = sheet['输入阻抗']
            detail_data_body = detail_data.body_data
            detail_data_body_col_2 = [float(row[1]) for row in detail_data_body]
            detail_data_head = detail_data.head_data

            range_str = ""
            # 取出第二列的最大值和最小值
            if len(detail_data_body) > 0:
                min_2 = round(min(detail_data_body_col_2), 0)
                max_2 = round(max(detail_data_body_col_2), 0)
                unit_2 = detail_data_head[1].split('[')[1].split(']')[0]
                range_str = "({}-{}{})".format(min_2, max_2, unit_2)

            unit_mhz = "MHz"

            col_worst = ColValue()
            seq_new += 0.1
            col_worst.seq = round(seq_new, 1)
            col_worst.item = '输入阻抗{}'.format(range_str)

            res_data.append(col_worst)
            # 第二行
            col_second = ColValue()
            col_second.seq = ""
            col_second.item = "上限最差值{}{}".format(row_data[0][2], unit_mhz)
            col_second.requirement = "最大{}".format(round(float(row_data[0][7]), 0))
            col_second.result = round(float(row_data[0][5]), 1)
            col_second.unit = row_data[0][0].split('[')[1].split(']')[0]
            col_second.assessment = 'P'
            res_data.append(col_second)

            # 第三行
            col_third = ColValue()
            col_third.seq = ""
            col_third.item = "下限最差值{}{}".format(row_data[0][2], unit_mhz)
            col_third.requirement = "最小{}".format(round(float(row_data[0][3]), 0))
            col_third.result = round(float(row_data[0][1]), 1)
            col_third.unit = row_data[0][0].split('[')[1].split(']')[0]
            col_third.assessment = 'P'
            res_data.append(col_third)

            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

        # 情况2  除了打印工作表里的数据外，还需要打印最差情形汇总工作表里的最差数据；（除了情况1和情况3以外的参数
        else:

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

                col_req = value.head_data[2]
                col_fill_str = "最大"
                if "下限" in col_req:
                    col_fill_str = "最小"

                # 获取结果列
                body_data = value.body_data
                for index, row in enumerate(body_data):
                    if seq_new_case:
                        col = ColValue()
                        # 保留一个小数
                        seq_new = seq_new + 0.1
                        seq_new_str = round(seq_new, 1)
                        col.item = item
                        col.seq = str(seq_new_str)

                        seq_new_case = False
                        res_data.append(col)

                    col = ColValue()

                    col.item = row[1] + ' ' + col_unit
                    col.unit = col_requirement
                    col.requirement = (col_fill_str if "最" not in row[2] else "") + row[2]
                    col.result = row[3]
                    col.assessment = 'P'
                    res_data.append(col)

            # 还需要获取最差情形，加到最后一行
            worst_data_now = sheet
            sheet_key = list(worst_data_now.keys())[0]
            worst_sheet = worst_data_now[sheet_key].body_data
            worst_head = worst_data_now[sheet_key].head_data
            worst_unit = worst_head[1].split('[')[1].split(']')[0]

            # 取出第二行为list
            worst_2 = [float(row[1]) for row in worst_sheet]
            # 最小值
            min_2 = round(min(worst_2), 0)
            # 最大值
            max_2 = round(max(worst_2), 0)

            # 判断是上限或下限
            # 下限用最小值和标准下限
            case_use = list(sheet.keys())
            head_data_case = sheet[case_use[0]].head_data[2]

            last_req = ""
            last_res = ""

            case_name = "最小"
            case_seq = 1  # 最小值的列序号
            case_seq_freq = 2
            case_std_seq = 3  # 标准下限的列序号
            if "上限" in head_data_case:
                case_seq = 5
                case_seq_freq = 6
                case_std_seq = 7
                case_name = "最大"

            # 取 最差情形汇总 的sheet
            i_ = excel_data.sheets.index("最差情形汇总")
            bad_sheet = excel_data.sheets_data[i_]

            bad_sheet_data = bad_sheet["最差情形汇总"].body_data

            # 取出对应行
            bad_sheet_data_row = [row for row in bad_sheet_data if sheet_key in row[0]]

            # 取出最差情形的值
            case_seq_value = bad_sheet_data_row[0][case_seq]
            case_seq_freq_value = bad_sheet_data_row[0][case_seq_freq]
            case_res_value = bad_sheet_data_row[0][case_std_seq]

            case_seq_value_str = str(round(float(case_seq_value), 1))
            case_res_value_str = str(round(float(case_res_value), 1))

            # 取单位
            case_unit = bad_sheet_data_row[0][0].split('[')[1].split(']')[0]

            # 组成行
            min_max_2 = "{}".format(min_2) + " - " + "{}".format(max_2) + worst_unit + "最差值" + f"({case_seq_freq_value})MHz"
            col_worst = ColValue()
            col_worst.seq = ""
            col_worst.item = min_max_2
            col_worst.requirement = case_name + case_res_value_str
            col_worst.result = case_seq_value_str
            col_worst.unit = case_unit
            col_worst.assessment = 'P'
            res_data.append(col_worst)


            # 一个sheet做完，空2行
            res_data.append(col_empty)
            res_data.append(col_empty)

    out_data = []
    for col in res_data:
        out_data.append([col.seq, col.item, col.unit, col.requirement, col.result, col.assessment])
    for i in out_data:
        print(i)
    return res_data


if __name__ == "__main__":
    file_path = './data/20240905-原始测试数据.xls'
    res = clean_data(file_path)

