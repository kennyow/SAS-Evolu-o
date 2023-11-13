def convert_to_importrange(link, sheet_name, range_str):
    # Extract the document ID from the link
    doc_id_start = link.find('/d/') + 3
    doc_id_end = link.find('/', doc_id_start)
    doc_id = link[doc_id_start:doc_id_end]

    # Construct the IMPORTRANGE formula
    formula = f'=IMPORTRANGE("{doc_id}"; "{sheet_name}!{range_str}")'

    return formula


def convert_to_importrange3(link, sheet_name1, range_str1, sheet_name2, range_str2, sheet_name3, range_str3 ):
    # Extract the document ID from the link
    doc_id_start = link.find('/d/') + 3
    doc_id_end = link.find('/', doc_id_start)
    doc_id = link[doc_id_start:doc_id_end]
    
    # Construct the IMPORTRANGE formula
    inicio = '=QUERY({'
    formula1 = f'IMPORTRANGE("{doc_id}"; "{sheet_name1}!{range_str1}");'
    formula2 = f'IMPORTRANGE("{doc_id}"; "{sheet_name2}!{range_str2}");'
    formula3 = f'IMPORTRANGE("{doc_id}"; "{sheet_name3}!{range_str3}")'
    end = '}; "SELECT Col1, Col2, Col3, Col4")'
    total = str(inicio+formula1+formula2+formula3+end)

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

#Student's Quantity

data_dict = {"6": {"6A":28, "6B":27, "6C":31, "6D":23}, 
             "7":{"7A":35, "7B": 38, "7C":35},
              "8":{"8A":32, "8B":38, "8C":35},
            "9":{"9A":40, "9B": 38, "9C":46},
            "1M":{"1AM":30, "1BM": 38, "1CM":29},
             "2M":{"2AM":48, "2BM": 48},
              "3M":{"3AM":33, "3BM":24}}



# Example usage
google_sheets_link = "https://docs.google.com/spreadsheets/d/1NDbPxtN8rqSKHFrdQ6PBml5nJgCH6FL4qReMMfZ2_6k/edit#gid=1340196971"
nome = 'Matemática'

turma = str(input("Qual ANO deseja gerar? "))

for ano in data_dict:
    target_sheet_name1 = "A2_6A_"+str(nome)+""
    target_range1 = "$A$7:$D$34"
    target_sheet_name2 = "A2_6B_"+str(nome)+""
    target_range2 = "$A$7:$D$33"
    target_sheet_name3 = "A2_6C_"+str(nome)+""
    target_range3 = "$A$7:$D$37"
    target_sheet_name5 = "A2_6D_"+str(nome)+""
    target_range5 = "$A$7:$D$29"
    target_sheet_name4 = "6ANO_"+(nome[:3]).upper()+""
    target_range4 = "$A$1:$E$126"

importrange_formula_1 = convert_to_importrange(google_sheets_link, target_sheet_name1, target_range1)
importrange_formula_3 = convert_to_importrange3(google_sheets_link, target_sheet_name1, target_range1, target_sheet_name2, target_range2, target_sheet_name3, target_range3)
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

