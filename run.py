import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from operation import Operations
from bilan_operation import bilan_operations

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

#regroupement des listes
def ReadAllExcel():
    arvida_lst  = ExcelfileToList("SAINT-DOMINIQUE")
    kenogami_lst  = ExcelfileToList("SAINTE-FAMILLE")
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
        if(item.eglise == "SAINT-DOMINIQUE"):
            data_cell = sheet.cell(row=x+indexToStart, column=9)
            data_cell.value = list[x].amount
        elif(item.eglise == "SAINTE-FAMILLE"):
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
    d4.value = "BUDGET année"
    d4.font = Font(bold=True,size=14)

    i4 = sheet.cell(row=4, column=9)
    i4.value = "SAINT-DOMINIQUE"
    i4.font = Font(bold=True)

    j4 = sheet.cell(row=4, column=10)
    j4.value = "SAINTE-FAMILLE"
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


#regroupement
def regroupement(list_operations):
    grouped_list = [bilan_operations(40100,"QUÊTES",0,0),
                    bilan_operations(40200,"CAPITATION",0,0),
                    bilan_operations(40300,"LUMINAIRE (CULTE)",0,0),
                    bilan_operations(40400,"CÉLÉBRATIONS",0,0),
                    bilan_operations(40500,"QUÊTES COMMANDÉES",0,0),
                    bilan_operations(40600,"DONS",0,0),
                    bilan_operations(40700,"PASTORALE",0,0),
                    bilan_operations(40800,"OBJETS DE REVENTE(FEUILLET)",0,0),
                    bilan_operations(40900,"EXTRAIT DES ACTES",0,0),
                    bilan_operations(41000,"ACTICITÉ DE FINANCEMENT",0,0),
                    bilan_operations(49000,"AUTRES PROVENANT DES",0,0),
                    bilan_operations(50100,"LOCATIONS",0,0),
                    bilan_operations(50200,"INTÉRETS",0,0),
                    bilan_operations(50300,"SUBVENTION",0,0),
                    bilan_operations(50400,"ristournes assurance",0,0),
                    bilan_operations(59000,"revenus autres",0,0),
                    bilan_operations(60100,"SALAIRE ET C.E.SACRISTAIN",0,0),
                    bilan_operations(60200,"MINISTÈRE",0,0),
                    bilan_operations(60300,"FRAIS DE VOYAGE",0,0),
                    bilan_operations(60400,"CÉLÉBRATIONS",0,0),
                    bilan_operations(60500,"FEUILLET PAROISSIAL",0,0),
                    bilan_operations(60600,"cultes",0,0),
                    bilan_operations(60700,"UNITÉ DES DEUX-RIVES",0,0),
                    bilan_operations(60800,"ANIMATION LITURGIQUE",0,0),
                    bilan_operations(60900,"ANIMATION PASTORALE",0,0),
                    bilan_operations(61000,"PART ÉGLISE",0,0),
                    bilan_operations(61100,"OBJETS DE REVENTE",0,0),
                    bilan_operations(61200,"CIMETIÈRE",0,0),
                    bilan_operations(61300,"Quêtes commandéée",0,0),
                    bilan_operations(61400,"TRIBUT DIOCÉSAIN",0,0),
                    bilan_operations(70100,"SALAIRES C.E.",0,0),
                    bilan_operations(70200,"DÉPENSES DE BUREAU",0,0),
                    bilan_operations(70300,"HONORAIRES ET CONTRATS",0,0),
                    bilan_operations(70400,"FORMATION",0,0),
                    bilan_operations(70500,"ADMINISTRATION/TPS et TVQ",0,0),
                    bilan_operations(70600,"CIVILITÉES",0,0),
                    bilan_operations(79000,"AUTRE DÉPENSES DE BUREAU",0,0),
                    bilan_operations(80100,"SALAIRE ET C.E. EMPLOYEUR",0,0),
                    bilan_operations(80200,"CHAUFFAGE",0,0),
                    bilan_operations(80300,"ÉLECTRICITÉ",0,0),
                    bilan_operations(80400,"ENTRETIEN INTÉRIEUR",0,0),
                    bilan_operations(80500,"ENTRETIEN EXTÉRIEUR",0,0),
                    bilan_operations(80600,"RÉPARATIONS MAJEURES",0,0),
                    bilan_operations(80700,"ASSURANCE",0,0),
                    bilan_operations(89000,"AUTRE DÉPENSES SUR BÂTISSES",0,0)
                    ]
    
    for operation in list_operations:
        for bilan in grouped_list:
            if(int(str(operation.account)[:3]) == int(str(bilan.account)[:3])):
                if(operation.eglise == "SAINT-DOMINIQUE"):
                    bilan.st_dominique_amount = bilan.st_dominique_amount+operation.amount
                elif(operation.eglise == "SAINTE-FAMILLE"):
                    bilan.sainte_famille_amount = bilan.sainte_famille_amount+operation.amount

    return 0



#prendre une capture écran de la premiere version avant de passer à la prochaine... pour éric
                
#debut programme
final_list = ReadAllExcel()
regroupement(final_list)