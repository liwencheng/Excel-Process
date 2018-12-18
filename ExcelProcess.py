# -*- coding:utf-8 -*-
import xlrd
import xlwt

new_workbook = xlwt.Workbook(encoding='utf-8')
new_booksheet = new_workbook.add_sheet('Open Orders', cell_overwrite_ok=True)

stylei = xlwt.XFStyle()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_LEFT
alignment.vert = xlwt.Alignment.VERT_CENTER
stylei.alignment = alignment
patterni = xlwt.Pattern()
patterni.pattern = 1
# 设置底纹的图案索引，1为实心，2为50%灰色，对应为excel文件单元格格式中填充中的图案样式
patterni.pattern_fore_colour = 50    # 设置底纹的前景色，对应为excel文件单元格格式中填充中的背景色
patterni.pattern_back_colour = 35   # 设置底纹的背景色，对应为excel文件单元格格式中填充中的图案颜色
stylei.pattern = patterni           # 为样式设置图案


styleMain = xlwt.XFStyle()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_LEFT
alignment.vert = xlwt.Alignment.VERT_CENTER
styleMain.alignment = alignment

#日期格式

dateFormat = xlwt.XFStyle()
dateFormat.num_format_str = 'yyyy/mm/dd'
dateFormat.alignment = alignment

new_workbook = xlwt.Workbook(encoding='utf-8')
new_booksheet = new_workbook.add_sheet('Open Orders', cell_overwrite_ok=True)

def copyExcel(source, des):
    workbook = xlrd.open_workbook(source)
    booksheet = workbook.sheet_by_name('Open Orders')
    first_row = booksheet.row_values(1)

    offset = 0;
    for j in range(len(first_row)):
        if(j == 0 or j == 4 or j == 12 or j== 13 or j== 14 or j == 17):
            continue
        new_booksheet.write(0, offset, first_row[j], stylei)
        offset +=1

    nrows = booksheet.nrows
    num = 1;


    for i in range(1,nrows):
        if(booksheet.row_values(i)[1] == des):
            offset = 0;
            for j in range(len(booksheet.row_values(i))):
                if(j == 0 or j == 4 or j == 12 or j== 13 or j== 14 or j == 17):
                    continue
                if(j == 8 or j == 15 or j == 21 or j == 20 or j ==24 or j ==25 or j == 26):
                    new_booksheet.write(num, offset,booksheet.row_values(i)[j],dateFormat)
                    offset +=1
                    continue
                new_booksheet.write(num, offset,booksheet.row_values(i)[j], styleMain)
                offset +=1
            num +=1

    for i in range(booksheet.ncols):
        new_booksheet.col(i).width = 4500

    new_workbook.save(des + '.xls')
    print "finish!"
    return;
