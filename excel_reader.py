import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from operation import Operations
from bilan_operation import bilan_operations

def ExcelfileToList(name_file):
    '''
    Returns list from a excel ( formated ).

            Parameters:
                    n (str): name of the file

            Returns:
                    binary_sum (str): Binary string of the sum of a and b
    '''
    file = name_file + ".xlsx"
    dataframe = openpyxl.load_workbook(file)
    
    # Define variables
    dataframe1 = dataframe.active
    list = []
    cpt=0
    temp_amount=0
    temp_name=""
    temp_no_account =0
    temp_type=0 # 0 = revenue, 1= depense
    string_list = ['Revenus :'] # filtre indésirables
    string_to_switch="Frais d'exploitation :" #mot pour switch entre revenues et dépsenses
    string_to_stop ="Total des frais"
    is_over = 0 #bool
    
    # Iteration dans les row
    for row in range(4, dataframe1.max_row):
        for col in dataframe1.iter_cols(1, 3): # prendre que les 2 premieres col des row
            if(col[row].value != string_to_stop and is_over == 0): 
                if(col[row].value not in string_list and col[row].value != None): #verifier si on a pas de data indésirable (filtre)
                    if(cpt==0):
                        temp_no_account = col[row].value
                        cpt = cpt+1
                    elif(cpt==1):
                            if(col[row].value == string_to_switch):    #switch entre les revenues et les dépenses
                                temp_type = 1
                            temp_name = col[row].value
                            cpt = cpt+1
                    else:
                        temp_amount = col[row].value
                        cpt =0
                        list.append(Operations(temp_no_account,temp_name,temp_type,temp_amount,name_file)) #ajout a la liste
            else:
                is_over=1
    return list

def ReadAllExcel():
    '''
    Read all excels, made and group list and return final list

            Parameters:

            Returns:
                    final_list (List[operation])
    '''
    arvida_lst  = ExcelfileToList("SAINT-DOMINIQUE")
    kenogami_lst  = ExcelfileToList("SAINTE-FAMILLE")
    final_lst = arvida_lst + kenogami_lst

    for x in final_lst:
        print(x.name+" - montant : "+str(x.amount))
    return final_lst