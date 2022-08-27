import xlwings as xw
from docxtpl import DocxTemplate
from docx2pdf import convert
import win32api
import os

def msg_box(message):
    wb = xw.Book.caller()
    win32api.MessageBox(wb.app.hwnd,message)

sheet = None
def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def save(last_row):
    path = os.path.dirname(os.path.realpath(__file__))+"\\"
    doc = DocxTemplate(path+"template.docx")
    no = int(getValue('B',last_row))
    context = { 
    'no' : no ,
    'date' : getValue('A',last_row) ,
    'seller' : getValue('C',last_row) ,
    'buyer' : getValue('D',last_row) ,
    'seller_phone' :int(getValue('E',last_row)) ,
    'buyer_phone' : int(getValue('F',last_row)) ,
    'seller_place' : getValue('G',last_row) ,
    'buyer_place' : getValue('H',last_row) ,
    'seller_id' : int(getValue('I',last_row)) ,
    'buyer_id' : int(getValue('J',last_row)) ,
    'alley' : getValue('O',last_row) ,
    'area' : getValue('N',last_row) ,
    'num' : getValue('M',last_row) ,
    'estate_t' : getValue('L',last_row) ,
    'estate' : getValue('K',last_row) ,
    'price' : getValue('R',last_row),
    'price' : getValue('R',last_row),
    'price_text' : getValue('S',last_row) ,
    'p_p' : getValue('T',last_row) ,
    'p_text' : getValue('U',last_row) ,
    'l_p' :  getValue('V',last_row),
    'l_text' : getValue('W',last_row) ,
    'undo' : getValue('X',last_row) ,
    'undo_text' : getValue('Y',last_row) ,
    'agree_d' : getValue('P',last_row) ,
    'eva_d' : getValue('Q',last_row) ,
    'witness1' : getValue('AA',last_row) ,
    'witness2' : getValue('AB',last_row) ,
    'unit' : getValue('AC',last_row) }
    doc.render(context)
    doc_path =path+"docs//doc_"+str(no)+".docx" 
    doc.save(doc_path)
    pdf_path = path+"pdfs//"+str(no)+'.pdf'
    convert(doc_path, pdf_path)
    msg_box('done')

@xw.func
def getLastRow():
    wb = xw.Book.caller()
    global sheet
    sheet = wb.sheets[0]
    last_row = str(wb.sheets[0].range('A' + str(sheet.cells.last_cell.row)).end('up').row)
    # sheet['A1'].value= sheet[last_row].value
    # msg_box(last_row)
    save(last_row)
    # return f"Hello {v,!"

def getValue(alpha,last_row):
    # msg_box(sheet)
    global sheet
    return sheet[alpha+last_row].value

if __name__ == "__main__":
    xw.Book("roobar.xlsm").set_mock_caller()
    main()