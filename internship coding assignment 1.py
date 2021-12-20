import fitz
import xlwt

def open_file(fname):

    doc = fitz.open(fname) 
    page = doc.load_page(11)
    text = page.get_text("blocks")
   #  put close file 
    
    my_xls = xlwt.Workbook(encoding='ascii', style_compression=0)
    my_sheet = my_xls.add_sheet("Paragraphs-Sheet", cell_overwrite_ok=True)
    
    row_no = 0
    
    for paragraph in text:
        
        if (paragraph[6] == 0):
            my_sheet.write(row_no,0,paragraph[4])
            row_no+=1
            
    my_xls.save(r"C:\Users\home\Documents\Python-Programs\results.xls")
    doc.close()
    return()


open_file(r"C:\Users\home\Documents\Python-Programs\keppel.pdf")
