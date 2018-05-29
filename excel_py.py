import xlrd
import numpy
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # print number of sheets
    print (book.nsheets)
 
    # print sheet names
    print (book.sheet_names())
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
 
    # read a row
    print (first_sheet.row_values(0))
 
    # read a cell
    cell = first_sheet.cell(0,0)
    print (cell)
    print (cell.value)
 
    # read a row slice
    print (first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2))


def get_all_correl():
    
    book = xlrd.open_workbook("data.xlsx")
    sheet = book.sheet_by_index(0)   
    all_cols = sheet.row_values(0)

    
    #for x in range(0,len(all_cols)):
#        print(x," -",all_cols[x] )
        
    all_cor = []
    all_header = []
    chosen_column = int(input("Choose column index: "))
    
    array_val_1 = sheet.col_values(chosen_column)
    chosen_header = array_val_1[0]
    array_val_1.pop(0)
    array_val_1 = alpha_to_numeric(array_val_1)
    for i in range(0,len(all_cols)):
        if(i != chosen_column and i !=0 and i!=3 and i<60):
            
            array_val_2 = sheet.col_values(i)
            header = array_val_2[0]
            array_val_2.pop(0)
            array_val_2 = alpha_to_numeric(array_val_2)
                
            tite = numpy.corrcoef(array_val_1,array_val_2)[1,0]
            all_cor.append(tite)
            all_header.append(header)

    for x in range(0, len(all_cor)):
        for y in range(0, len(all_cor)):
            
            if(all_cor[y] > all_cor[x]):
                num = all_cor[x]
                header = all_header[x]

                all_cor[x] = all_cor[y]
                all_cor[y] = num
                
                all_header[x] = all_header[y]
                all_header[y] = header

    list.reverse(all_header)
    list.reverse(all_cor)
    
 #   print_all(all_header, all_cor)
    print_top_ten(all_header,all_cor)
    print("-----")
    print_bot_ten(all_header,all_cor)
    print("-----")
def print_all(all_header, all_cor):
    for i in range(0, len(all_header)):
        print(all_header[i],": " ,all_cor[i])
        
def print_top_ten(all_header,all_cor):
    for i in range(0, 9):
        print(all_header[i],": " ,all_cor[i])
def print_bot_ten(all_header,all_cor):
    for i in range(-9, 0):
        print(all_header[i],": " ,all_cor[i])
def alpha_to_numeric(array):
    for i in range(0,len(array)):
        if(array[i] =="na"):
            array[i] = 0
    return array

while (True):
    get_all_correl()
    cont = input("press [n] to stop")
    if (cont=="n"):
        print("thank")
        break
