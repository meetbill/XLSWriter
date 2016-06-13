#!/usr/bin/python
#coding=utf8
# xlswriter.py
"""
# Author: Bill
# Created Time : 2016年02月23日 星期二 16时34分11秒

# File Name: w.py
# Description:

"""
import mylib.xlwt

class XLSWriter(object):
    """A XLS writer that produces XLS files from unicode data.
    """
    def __init__(self, file, encoding='utf-8'):
        # must specify the encoding of the input data, utf-8 default.
        self.file = file
        self.encoding = encoding
        self.wbk = mylib.xlwt.Workbook()
        self.sheets = {}
        
    def add_image(self,bmp_name='',x='',y='',sheet_name='sheet'):
        if sheet_name not in self.sheets:
            # Create if does not exist
            self.create_sheet(sheet_name)
        self.sheets[sheet_name]['sheet'].insert_bitmap('python.bmp',\
                                                       x,y,0,0,scale_x=0.1,\
                                                       scale_y=0.1)
        self.sheets[sheet_name]['rows'] += 1
    def create_sheet(self, sheet_name='sheet'):
        """Create new sheet
        """
        if sheet_name in self.sheets:
            sheet_index = self.sheets[sheet_name]['index'] + 1
        else:
            sheet_index = 0
            self.sheets[sheet_name] = {'header': []}
        self.sheets[sheet_name]['index'] = sheet_index
        self.sheets[sheet_name]['sheet'] = self.wbk.add_sheet('%s%s' % (sheet_name, sheet_index if sheet_index else ''), cell_overwrite_ok=True)
        self.sheets[sheet_name]['rows'] = 0

    def cell(self, s):
        if isinstance(s, basestring):
            if not isinstance(s, unicode):
                s = s.decode(self.encoding)
        elif s is None:
            s = ''
        else:
            s = str(s)
        return s

    def writerow(self, row, sheet_name='sheet',border=False):
        if border:
            borders = mylib.xlwt.Borders()
            borders.left = 1
            borders.right = 1
            borders.top = 1
            borders.bottom = 1
            borders.bottom_colour=0x3A    
         
            style = mylib.xlwt.XFStyle()
            style.borders = borders 
          
        #sheet.write(0, 0, 'Firstname',style)
        if sheet_name not in self.sheets:
            # Create if does not exist
            self.create_sheet(sheet_name)
    
        if self.sheets[sheet_name]['rows'] == 0:
            self.sheets[sheet_name]['header'] = row

        if self.sheets[sheet_name]['rows'] >= 65534:
            self.save()
            # create new sheet to avoid being greater than 65535 lines
            self.create_sheet(sheet_name)
            if self.sheets[sheet_name]['header']:
                self.writerow(self.sheets[sheet_name]['header'], sheet_name)
        for ci, col in enumerate(row):
            if border:
                self.sheets[sheet_name]['sheet'].write(self.sheets[sheet_name]['rows'],ci,\
                                                   self.cell(col) if type(col) != mylib.xlwt.ExcelFormula.Formula else col,\
                                                   style)
            else:
                self.sheets[sheet_name]['sheet'].write(self.sheets[sheet_name]['rows'],ci,\
                                                   self.cell(col) if type(col) != mylib.xlwt.ExcelFormula.Formula else col)
            #print self.sheets[sheet_name]['rows'],ci, self.cell(col) if type(col) != lib.xlwt.ExcelFormula.Formula else col
        self.sheets[sheet_name]['rows'] += 1
            
    def writerows(self, rows, sheet_name='sheet'):
        for row in rows:
            self.writerow(row, sheet_name)

    def save(self):
        self.wbk.save(self.file)
        
if __name__ == '__main__':
    # test
    xlswriter = XLSWriter('ceshi.xls')
    xlswriter.add_image("python.bmg",0,0,sheet_name=u"基本信息")
    xlswriter.writerow(['姓名', '年龄', '电话', 'QQ'],sheet_name=u'基本信息',border=True)
    xlswriter.writerow(['张三', '30', '12345678910', '123456789'], sheet_name=u'基本信息',border=True)
    xlswriter.writerow(['王五', '30', '13512345678', '123456789'], sheet_name=u'基本信息',border=True)
    
    xlswriter.writerow(['检测项', '检测命令', '值','基准','状态'],sheet_name=u'服务器器状态')
    xlswriter.writerow(["磁盘空间", "df -lP | grep -e '/$' | awk '{print $5}'","20%","%85","OK"], sheet_name=u'服务器器状态')
    # don't forget to save data to disk
    xlswriter.save()
    print 'finished.'
