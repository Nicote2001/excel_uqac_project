import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from operation import Operations

# fonction de lecture de fichier pour transformer en liste
def ExcelfileToList(name_file):
    file = name_file + ".xlsx"
    dataframe = openpyxl.load_workbook(file)
    
    # Define variables
    dataframe1 = dataframe.active
    list = []
    cpt=0
    temp_amount=0
    temp_name=""
    temp_type=0 # 0 = revenue, 1= depense
    string_list = ['Revenus :'] # filtre indésirables
    string_to_switch="Frais d'exploitation :" #mot pour switch entre revenues et dépsenses
    string_to_stop ="Total des frais"
    is_over = 0 #bool
    
    # Iteration dans les row
    for row in range(4, dataframe1.max_row):
        for col in dataframe1.iter_cols(1, 2): # prendre que les 2 premieres col des row
            if(col[row].value != string_to_stop and is_over == 0): 
                if(col[row].value not in string_list and col[row].value != None): #verifier si on a pas de data indésirable (filtre)
                    if(cpt==0):
                        if(col[row].value == string_to_switch):    #switch entre les revenues et les dépenses
                            temp_type = 1
                        else:
                            temp_name = col[row].value
                            cpt = cpt+1
                    else:
                        temp_amount = col[row].value
                        cpt =0
                        list.append(Operations(temp_name,temp_type,temp_amount,name_file)) #ajout a la liste
            else:
                is_over=1
    return list

#regroupement des listes
def ReadAllExcel():
    arvida_lst  = ExcelfileToList("arvida")
    kenogami_lst  = ExcelfileToList("kenogami")
    final_lst = arvida_lst + kenogami_lst

    for x in final_lst:
        print(x.name+" - montant : "+str(x.amount))
    return final_lst

def SplitListByType(type, list):
    list_return = []
    for x in list:
        if(x.type == type):
            list_return.append(x)
    return list_return

#on passe la list a trier dans le excel, la feuille Excel et l'index de commencement 
def WriteOperation(list, sheet, indexToStart):
    for x, item in enumerate(list):
        data_cell = sheet.cell(row=x+indexToStart, column=6)
        data_cell.value = list[x].name
        if(item.eglise == "arvida"):
            data_cell = sheet.cell(row=x+indexToStart, column=9)
            data_cell.value = list[x].amount
        elif(item.eglise == "kenogami"):
            data_cell = sheet.cell(row=x+indexToStart, column=10)
            data_cell.value = list[x].amount

def WriteExcel(list_operation):
    
    #open ExcelSheetWorker
    wb = openpyxl.Workbook()
    sheet = wb.active  

    #set default col dimensions
    sheet.column_dimensions['F'].width = 40
    sheet.column_dimensions['I'].width = 20
    sheet.column_dimensions['J'].width = 20
    sheet.column_dimensions['K'].width = 20
    sheet.column_dimensions['L'].width = 20

    #split list by types
    revenue_list = SplitListByType(0,list_operation)
    depense_list = SplitListByType(1,list_operation)


    #titre en gras et en gros
    g2 = sheet.cell(row=2, column=7)
    g2.value = "REGROUPEMENT DES PAROISSES"
    g2.font = Font(bold=True,size=18)
    g2.alignment = Alignment(horizontal="center")

    #chiffre en haut des parroises
    for x in range(1,5):
        temp_cell = sheet.cell(row=3, column=x+8)
        temp_cell.value = x
        temp_cell. alignment = Alignment(horizontal="center")

    #date et parroisses
    d4 = sheet.cell(row=4, column=4)
    d4.value = "RÉSULTAT date-mois-jours"
    d4.font = Font(bold=True,size=14)

    i4 = sheet.cell(row=4, column=9)
    i4.value = "Arvida"
    i4.font = Font(bold=True)

    j4 = sheet.cell(row=4, column=10)
    j4.value = "Kenogami"
    j4.font = Font(bold=True)

    #revenus
    f5 = sheet.cell(row=5, column=6)
    f5.value = "REVENUS"
    f5.font = Font(bold=True,size=14)

    #écriture des revenues
    WriteOperation(revenue_list, sheet, 6)

    #revenus
    dep = sheet.cell(row=len(revenue_list)+7, column=6)
    dep.value = "DÉPENSES"
    dep.font = Font(bold=True,size=14)

    #écriture depenses
    WriteOperation(depense_list, sheet, len(revenue_list)+8)
    
    wb.save("demo.xlsx")


            
                
#debut programme
final_list = ReadAllExcel()
WriteExcel(final_list)