import xlwt
from datetime import datetime
from xlwt import easyxf

tl = easyxf('border: left thick, top thick')
t = easyxf('border: top thick')
tr = easyxf('border: right thick, top thick')
r = easyxf('border: right thick')
br = easyxf('border: right thick, bottom thick')
b = easyxf('border: bottom thick')
bl = easyxf('border: left thick, bottom thick')
l = easyxf('border: left thick')
# ws.row(8).write(0,'',Style.easyxf('pattern: pattern solid, fore_colour green;'))

tlb = easyxf('border: left double, top double, bottom double')
trb = easyxf('border: right double, top double, bottom double')
bb = easyxf('border: bottom double')
tt = easyxf('border: top double')
ttbb = easyxf('border: top double, bottom double')

font0 = xlwt.Font()
# font0.name = 'Times New Roman'
font0.name = 'Arial'
font0.colour_index = 2
font0.bold = True
font0.italic = False
font0.escapement = False
font0.shadow = False

style0 = xlwt.XFStyle()
style0.font = font0

style1 = xlwt.XFStyle()
style1.num_format_str = 'D-MMM-YY'

wb = xlwt.Workbook() 
# ws = wb.add_sheet('A Test Sheet')
# wss = wb.add_sheet('Next one')
# wsss = wb.add_sheet('Another one')
sheet4 = wb.add_sheet('New Four')

# ws.write(0, 0, 'Test', style0)
# ws.write(1, 0, datetime.now(), style1)
# ws.write(2, 0, 1, style0)
# ws.write(2, 1, 1, style0)
# ws.write(2, 2, xlwt.Formula("A3+B3"))

# wss.write(0, 0, 'Next', style0)
# wss.write(1, 0, datetime.now(), style1)
# wss.write(2, 0, 1, style0)
# wss.write(2, 1, 1, style0)
# wss.write(2, 2, xlwt.Formula("log(A3+B3)"))

# wsss.write(0, 0, 'Another one', style=tl)
# wsss.write(0, 1, '', style=t)
# wsss.write(0, 2, style=tr)
# wsss.write(1, 0, datetime.now(), style=l)
# wsss.write(1, 2, style=r)
# wsss.write(2, 0, 1, style=bl)
# wsss.write(2, 1, 1, style=b)
# text = 'You may write anything you would like to write here'
# wsss.write(2, 2, text, style=br)

# Индекс столбца и строки ячейки, для которой необходимо установить ширину)
# text_length = len(text)
# wsss.col(2).width = 230 * (text_length + 1)


sheet4.write(1, 1, 'Today,', style=tlb)
sheet4.write(1, 2, datetime.now(), style1)
sheet4.write(0, 2, '', style=bb)
sheet4.write(2, 2, '', style=tt)
text_a = "I'm"
text_b = 'Filatov V. Vladimir,'
text_c = 'The King! :)'
sheet4.write(1, 3, text_a, style=ttbb)
sheet4.write(1, 4, text_b, style=ttbb)
sheet4.write(1, 5, text_c, style=trb)
text_length_a = len(text_a)
sheet4.col(3).width = 230 * (text_length_a + 1)
text_length_b = len(text_b)
sheet4.col(4).width = 230 * (text_length_b + 1)
text_length_c = len(text_c)
sheet4.col(5).width = 230 * (text_length_c + 1)

wb.save('excel_write_VV.xls')
