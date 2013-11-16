#!/usr/bin/python
# -*- coding: utf-8 -*-
import os, sys
import xlrd
import xlsxwriter

def main(argv):
    filename = sys.argv[1]
    sum_by = sys.argv[2]
    if sum_by == 'sku':
	sum_by_sku(filename)
    elif sum_by == 'brand':
        sum_by_brand(filename)
    elif sum_by == 'customer':
        sum_by_customer(filename)
    elif sum_by == 'address':
	sum_by_address(filename)

def sum_by_sku(filename): 
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('sale_by_sku.xlsx')
    worksheet = workbook.add_worksheet()
    wb = xlrd.open_workbook(filename)
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

    worksheet.write(0, 0, 'SEGMENT1')
    worksheet.write(0, 1, 'QUANTITY')
    worksheet.write(0, 2, 'NET_AMOUNT')
    row = 1
    for code in sorted(item_code_q.keys()):
        worksheet.write(row, col, code)
        worksheet.write(row, col + 1, item_code_q[code])
        worksheet.write(row, col + 2, item_code_n[code])
        # print code, '   ', item_code_q[code], ' ', item_code_n[code]
        total_q += item_code_q[code]
        total_net += item_code_n[code]
        row += 1

    # print total_q, ' ', total_net
    worksheet.write(row, 0, 'TOTAL')
    worksheet.write(row, col + 1, total_q)
    worksheet.write(row, col + 2, total_net)
    workbook.close()

def sum_by_brand(filename):
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)
    
    workbook = xlsxwriter.Workbook('sum_by_brand.xlsx')
    worksheet = workbook.add_worksheet()

    count = 1
    items_key_by_customer = {}
    for rownum in range(sh.nrows - 2):
        key_customer_name =  sh.cell(count,6).value
        key_code_name =  sh.cell(count,1).value
        key_customer = sh.cell(count,7).value
        key_value = sh.cell(count,10).value
        key_value_net = sh.cell(count,13).value
        if key_customer in items_key_by_customer:
            if key_code_name in items_key_by_customer[key_customer]:
                items_key_by_customer[key_customer][key_code_name][0] += key_value
                items_key_by_customer[key_customer][key_code_name][1] += key_value_net
            else:
                items_key_by_customer[key_customer][key_code_name] = [key_value, key_value_net, key_customer_name]
        else:
            items_key_by_customer[key_customer] = {}
            items_key_by_customer[key_customer][key_code_name] = [key_value, key_value_net, key_customer_name]
     
	count += 1

    worksheet.write(0, 0, 'CUSTOMER')
    worksheet.write(0, 1, 'CUSTOMER_NAME')
    worksheet.write(0, 2, 'ITEM_NAME_NAME')
    worksheet.write(0, 3, 'QUANTITY')
    worksheet.write(0, 4, 'NET_AMOUNT')
    row = 1
    col = 0
    for customer_code in items_key_by_customer.keys():
        for item_code_name in items_key_by_customer[customer_code].keys():
	    worksheet.write(row, 0, customer_code)
	    worksheet.write(row, 1, items_key_by_customer[customer_code][item_code_name][2])
	    worksheet.write(row, 2, item_code_name)
	    worksheet.write(row, 3, items_key_by_customer[customer_code][item_code_name][0])
	    worksheet.write(row, 4, items_key_by_customer[customer_code][item_code_name][1])
	    row += 1
	worksheet.write(row, 0, '')
	row += 1

    workbook.close()
    
def sum_by_customer(filename):
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)
    
    workbook = xlsxwriter.Workbook('sum_by_customer.xlsx')
    worksheet = workbook.add_worksheet()

    count = 1
    items_key_by_customer = {}
    for rownum in range(sh.nrows - 2):
        key_item_code =  sh.cell(count,0).value
        key_code_name =  sh.cell(count,1).value
	customer_name = sh.cell(count,6).value
        key_customer = sh.cell(count,7).value
        key_value = sh.cell(count,10).value
        key_value_net = sh.cell(count,13).value
        if key_customer in items_key_by_customer:
            if key_item_code in items_key_by_customer[key_customer]:
                items_key_by_customer[key_customer][key_item_code][0] += key_value
                items_key_by_customer[key_customer][key_item_code][1] += key_value_net
            else:
                items_key_by_customer[key_customer][key_item_code] = [key_value, key_value_net, key_code_name, customer_name]
        else:
            items_key_by_customer[key_customer] = {}
            items_key_by_customer[key_customer][key_item_code] = [key_value, key_value_net, key_code_name, customer_name]
     
	count += 1

    worksheet.write(0, 0, 'CUSTOMER')
    worksheet.write(0, 1, 'CUSTOMER_NAME')
    worksheet.write(0, 2, 'ITEM_NAME')
    worksheet.write(0, 3, 'ITEM_CODE')
    worksheet.write(0, 4, 'QUANTITY')
    worksheet.write(0, 5, 'NET_AMOUNT')
    row = 1
    col = 0
    for customer_code in items_key_by_customer.keys():
        for item_code_name in sorted(items_key_by_customer[customer_code].keys()):
	    worksheet.write(row, 0, customer_code)
	    worksheet.write(row, 1, items_key_by_customer[customer_code][item_code_name][3])
	    worksheet.write(row, 2, items_key_by_customer[customer_code][item_code_name][2])
	    worksheet.write(row, 3, item_code_name)
	    worksheet.write(row, 4, items_key_by_customer[customer_code][item_code_name][0])
	    worksheet.write(row, 5, items_key_by_customer[customer_code][item_code_name][1])
	    row += 1
	worksheet.write(row, 0, '')
	row += 1

    workbook.close()

def sum_by_address(filename):
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)
    
    workbook = xlsxwriter.Workbook('sum_by_address.xlsx')
    worksheet = workbook.add_worksheet()

    count = 1
    items_key_by_customer = {}
    for rownum in range(sh.nrows - 2):
        key_item_code = sh.cell(count,0).value
        key_code_name =  sh.cell(count,1).value
        key_customer = sh.cell(count,7).value
        key_customer_name =  sh.cell(count,9).value
        key_value = sh.cell(count,10).value
        key_value_net = sh.cell(count,13).value
        if key_customer in items_key_by_customer:
            if key_customer_name in items_key_by_customer[key_customer]:
		if key_item_code in items_key_by_customer[key_customer][key_customer_name]:
                    items_key_by_customer[key_customer][key_customer_name][key_item_code][0] += key_value
                    items_key_by_customer[key_customer][key_customer_name][key_item_code][1] += key_value_net
		else:
		    items_key_by_customer[key_customer][key_customer_name][key_item_code] = [key_value, key_value_net, key_code_name]

            else:
		items_key_by_customer[key_customer][key_customer_name] = {}
                items_key_by_customer[key_customer][key_customer_name][key_item_code] = [key_value, key_value_net, key_code_name]

        else:
            items_key_by_customer[key_customer] = {}
            items_key_by_customer[key_customer][key_customer_name] = {}
	    items_key_by_customer[key_customer][key_customer_name][key_item_code] = [key_value, key_value_net, key_code_name]
     
	count += 1

    worksheet.write(0, 0, 'CUSTOMER')
    worksheet.write(0, 1, 'CUSTOMER_NAME')
    worksheet.write(0, 2, 'PRODUCT')
    worksheet.write(0, 3, 'ITEM_NAME')
    worksheet.write(0, 4, 'QUANTITY')
    worksheet.write(0, 5, 'NET_AMOUNT')
    row = 1
    col = 0
    for customer_code in items_key_by_customer.keys():
        for item_code_name in items_key_by_customer[customer_code].keys():
	    for item_code in items_key_by_customer[customer_code][item_code_name].keys():
		    worksheet.write(row, 0, customer_code)
		    worksheet.write(row, 1, item_code_name)
		    worksheet.write(row, 2, items_key_by_customer[customer_code][item_code_name][item_code][2])
		    worksheet.write(row, 3, item_code)
		    worksheet.write(row, 4, items_key_by_customer[customer_code][item_code_name][item_code][0])
		    worksheet.write(row, 5, items_key_by_customer[customer_code][item_code_name][item_code][1])
	    	    row += 1

	    row += 1
	
	worksheet.write(row, 0, '')
	row += 1

if __name__ == "__main__":
    main(sys.argv[1:2])
    print 'Done!!'
