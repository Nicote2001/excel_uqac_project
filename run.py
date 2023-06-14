import excel_reader, excel_writer



def val(sheet,x, y):
    return sheet.cell(row=x, column=y).value


#prendre une capture écran de la premiere version avant de passer à la prochaine... pour éric
                
#debut programme
final_list = excel_reader.ReadAllExcel()
bilan_list = excel_writer.regroupement(final_list)
excel_writer.WriteExcel(bilan_list)