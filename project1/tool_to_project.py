#!/usr/bin/env python3
# coding: utf-8


import shutil
import openpyxl


def main():
    shutil.copy("./template.xlsx", "./dust_test_data.xlsx")
    src = openpyxl.load_workbook('../source_test_data.xlsx')
    dst = openpyxl.load_workbook('./dust_test_data.xlsx')

    """ src.active """
    src.template = False
    """ dst.active """
    dst.template = False


    print('[S]SOURCE-INFO')
    print(type(src))
    print(src.sheetnames)
    print(src[src.sheetnames[0]].max_row)
    print(src[src.sheetnames[0]].max_column)
    print('[E]SOURCE-INFO')
    print('[S]DUST-INFO')
    print(type(dst))
    print(dst.sheetnames)
    print(dst[dst.sheetnames[0]].max_row)
    print(dst[dst.sheetnames[0]].max_column)
    print('[E]DUST-INFO')


    srcSheet = src[src.sheetnames[0]]
    dstSheet = dst[dst.sheetnames[0]]

    for row in range(1,srcSheet.max_row):
        print(type(srcSheet.cell(row=row, column=1).value))
        print(srcSheet.cell(row=row, column=1).value)

        value = ''
        for column in range(2, 5):
            cell = srcSheet.cell(row=row, column=column)
            if (cell.value != None):
                value += cell.value
                value += '\n'
        dstSheet.cell(row=row, column=1).value = value

        value = ''
        for column in range(6, 9):
            cell = srcSheet.cell(row=row, column=column)
            if (cell.value != None):
                value += cell.value
                value += '\n'
        dstSheet.cell(row=row, column=2).value = value
        
        value = ''
        for column in range(10, 13):
            cell = srcSheet.cell(row=row, column=column)
            if (cell.value != None):
                value += cell.value
                value += '\n'
        dstSheet.cell(row=row, column=3).value = value
    
    dst.save('dust_test_data.xlsx')


if __name__ == '__main__':
    print('[S]tool_to_project')
    main()
    print('[E]tool_to_project')