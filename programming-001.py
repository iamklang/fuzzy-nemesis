#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import xlrd
import xlsxwriter

def main(argv):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('sale_by_sku.xlsx')
    worksheet = workbook.add_worksheet()
    xls_path = sys.argv[1]
    wb = xlrd.open_workbook(xls_path)
    sh = wb.sheet_by_index(0)
    row = 0
    col = 0
    total_q = 0
    total_net = 0
    count = 1
    item_code_q = {}
    item_code_n = {}
    for rownum in range(sh.nrows - 2):
        key_code =  sh.cell(count,0).value
        key_value = sh.cell(count,10).value
        key_value_net = sh.cell(count,13).value
    	if key_code in item_code_q:
		key_value += item_code_q[key_code]
		item_code_q[key_code] = key_value
    	else:
		item_code_q[key_code] = key_value

	if key_code in item_code_n:
		key_value_net += item_code_n[key_code]
		item_code_n[key_code] = key_value_net
    	else:
		item_code_n[key_code] = key_value_net

        count += 1

    for code in sorted(item_code_q.keys()):
        worksheet.write(row, col, code)
        worksheet.write(row, col + 1, item_code_q[code])
        worksheet.write(row, col + 2, item_code_n[code])
        print code, '   ', item_code_q[code], ' ', item_code_n[code]
        total_q += item_code_q[code]
        total_net += item_code_n[code]
        row += 1

    print total_q, ' ', total_net
    worksheet.write(row, col + 1, total_q)
    worksheet.write(row, col + 2, total_net)
    workbook.close()

if __name__ == "__main__":
    main(sys.argv[1:])

