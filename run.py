import excel_reader, excel_writer

#debut programme
final_list = excel_reader.ReadAllExcel()
bilan_list = excel_writer.regroupement(final_list)
excel_writer.WriteExcel(bilan_list)