# coding: utf-8
# xlswriter.py
 
# http://pypi.python.org/pypi/xlwt
import xlwt
import sys
if sys.getdefaultencoding() != 'utf-8':
    reload(sys)
    sys.setdefaultencoding('utf-8')

class XLSWriter(object):
    """A XLS writer that produces XLS files from unicode data.
    """
    def __init__(self, file, encoding='utf-8'):
        # must specify the encoding of the input data, utf-8 default.
        self.file = file
        self.encoding = encoding
        self.wbk = xlwt.Workbook()
        self.sheets = {}
        
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
 
    def writerow(self, row,xlsstyle, sheet_name='sheet'):
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
		    #self.sheets[sheet_name]['sheet'].col(col).width=0x0d00
            self.sheets[sheet_name]['sheet'].write(self.sheets[sheet_name]['rows'], ci, self.cell(col) if type(col) != xlwt.ExcelFormula.Formula else col,xlsstyle)
        self.sheets[sheet_name]['rows'] += 1
            
    def writerows(self, rows,style, sheet_name='sheet'):
        for row in rows:
            self.writerow(row,style, sheet_name)
 
    def save(self):
        self.wbk.save(self.file)
        
if __name__ == '__main__':
    # test
	
    reload(sys)
    sys.setdefaultencoding('utf8')
    xlswriter = XLSWriter(u'陕西.xls')
 
    ft=xlwt.Font()
    ft.height =0x00C8 
    ft.bold = True
 
    ft1=xlwt.Font()
    ft1.bold=False
 
    style0=xlwt.XFStyle()
    style0.font=ft
 
    style1=xlwt.XFStyle()
    style1.font=ft1
 
    xlswriter.writerow(['姓名', '年龄', '电话', 'QQ'], style0,sheet_name=u'基本信息')
    xlswriter.writerow(['张三', '30', '13512345678', '123456789'],style1, sheet_name=u'基本信息')
    
    xlswriter.writerow(['学校', '获得学位', '取得学位时间'], style0,sheet_name=u'学习经历')
    xlswriter.writerow(['西安电子科技大学', '学士', '2009'],style1, sheet_name=u'学习经历')
    xlswriter.writerow(['西安电子科技大学', '硕士', '2012'], style1,sheet_name=u'学习经历')
    
    xlswriter.writerow(['王五', '30', '13512345678', '123456789'],style1, sheet_name=u'基本信息')
    # don't forget to save data to disk
    xlswriter.save()
    print 'finished.'
