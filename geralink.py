
#GERA LINKS PARA O GOOGLE PLANILHAS


def convert_to_importrange(link, sheet_name, range_str):
    # Extract the document ID from the link
    doc_id_start = link.find('/d/') + 3
    doc_id_end = link.find('/', doc_id_start)
    doc_id = link[doc_id_start:doc_id_end]

    # Construct the IMPORTRANGE formula
    formula = f'=IMPORTRANGE("{doc_id}"; "{sheet_name}!{range_str}")'

    return formula


def convert_to_importrange3(link, sheet_name1, range_str1, sheet_name2, range_str2, sheet_name3, range_str3 ):#dois últimos 6Dsheet_name5, range_str5
    # Extract the document ID from the link
    doc_id_start = link.find('/d/') + 3
    doc_id_end = link.find('/', doc_id_start)
    doc_id = link[doc_id_start:doc_id_end]
    
    # Construct the IMPORTRANGE formula
    inicio = '=QUERY({'
    formula1 = f'IMPORTRANGE("{doc_id}"; "{sheet_name1}!{range_str1}");'
    formula2 = f'IMPORTRANGE("{doc_id}"; "{sheet_name2}!{range_str2}");'
    formula3 = f'IMPORTRANGE("{doc_id}"; "{sheet_name3}!{range_str3}");'
    #formula5 = f'IMPORTRANGE("{doc_id}"; "{sheet_name5}!{range_str5}")'
    end = '}; "SELECT Col1, Col2, Col3, Col4")'
    total = str(inicio+formula1+formula2+formula3+end)#formula5+

    return total


def convert_to_importrange_single(link, sheet_name, range_str):
    # Extract the document ID from the link
    doc_id_start = link.find('/d/') + 3
    doc_id_end = link.find('/', doc_id_start)
    doc_id = link[doc_id_start:doc_id_end]

    # Construct the IMPORTRANGE formula
    inicio = '=QUERY({'
    formula = f'IMPORTRANGE("{doc_id}"; "{sheet_name}!{range_str}")'
    end = '}; "SELECT Col2, Col3, Col4, Col1, Col5")'
    total = str(inicio+formula+end)
    return total


# Example usage
google_sheets_link = 'https://docs.google.com/spreadsheets/d/1cXLT_TNqfp40X0B9O9qHtW9-fZIW-mC7opAwUGF7hM8/edit#gid=1355969042'
nome = 'PROTEX'
target_sheet_name1 = "A1_9A_"+str(nome)+""
target_range1 = "$A$7:$D$47"
target_sheet_name2 = "A1_9B_"+str(nome)+""
target_range2 = "$A$7:$D$45"
target_sheet_name3 = "A1_9C_"+str(nome)+""
target_range3 = "$A$7:$D$51"
'''target_sheet_name5 = "A2_6D_"+str(nome)+""
target_range5 = "$A$7:$D$30"'''
target_sheet_name4 = "9ANO_"+(nome[:3]).upper()+""
target_range4 = "$A$1:$E$129"

importrange_formula_1 = convert_to_importrange(google_sheets_link, target_sheet_name1, target_range1)
importrange_formula_3 = convert_to_importrange3(google_sheets_link, target_sheet_name1, target_range1, target_sheet_name2, target_range2, target_sheet_name3, target_range3) #target_sheet_name5, target_range5target5 é para o 6D
importrange_formula_s = convert_to_importrange_single(google_sheets_link, target_sheet_name4, target_range4)

print( )
print("\033[33m*\033[0m"*140)
print('\033[41mCÓDIGO DE '+str(nome)+'\033[0m')
print(importrange_formula_1)
print("-"*140)
print(importrange_formula_3)
print("-"*140)
print(importrange_formula_s)
print("\033[33m*\033[0m"*140)



