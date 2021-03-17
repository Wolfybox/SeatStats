import xlrd


def check_stu(stu_info_dir, spoc_dir):
    stu_data = xlrd.open_workbook(stu_info_dir)
    stu_sheet = stu_data.sheet_by_index(0)
    print(f'Total Students: {stu_sheet.nrows}')
    target_stu_list = []
    for row_index in range(stu_sheet.nrows):
        row_ele_list = stu_sheet.row_values(row_index)
        row_id, stu_id, stu_name, stu_acade, _, stu_class, _, _, ta_name = row_ele_list
        target_stu_list.append(stu_name)

    spoc_data = xlrd.open_workbook(spoc_dir)
    spoc_sheet = spoc_data.sheet_by_index(0)
    print(f'Total Students in Spoc: {spoc_sheet.nrows}')
    spoc_stu_list = []
    for row_index in range(1, spoc_sheet.nrows):
        row_ele_list = spoc_sheet.row_values(row_index)
        nickname, real_name, stu_no, college, group_name = row_ele_list
        spoc_stu_list.append(real_name)

    miss_names = []
    for name in target_stu_list:
        if name not in spoc_stu_list:
            miss_names.append(name)
    print(f'missing students: {miss_names}')
    print(f'Total:{len(miss_names)}')


if __name__ == '__main__':
    check_stu(stu_info_dir='我负责的同学.xlsx', spoc_dir='20210220-202103012021春大学计算机SPOC选课名单.xls')
