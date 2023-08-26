import openpyxl

wb = openpyxl.reader.excel.load_workbook(filename="f_3_08_2.xlsx")
wb.active =0
sheet = wb.active

wb1 = openpyxl.reader.excel.load_workbook(filename="f_4_08.xlsx")
wb1.active =0
sheet1 = wb1.active


def init_exel(i:int):
    if (i == 285): return

    old_table_name = sheet['A'+str(i)].value
    new_table_name = sheet1['A'+str(i)].value

    old_view_doc = sheet['F'+str(i)].value
    new_view_doc = sheet1['F'+str(i)].value

    direction = sheet['C'+str(i)].value
    if (old_table_name != new_table_name):
        print(i)
        print(old_table_name)
        print(new_table_name)
        return
    else:
        if (old_view_doc != new_view_doc):
            if (old_view_doc == None):
                old_view_doc = "-"
            if (new_view_doc == None):
                new_view_doc = "-"
            print(old_table_name+ " dir:"+direction+" error:"+old_view_doc+"(old)  "+new_view_doc+"(new)")
            
    init_exel(i+1)


if __name__ == '__main__':
    init_exel(1)
