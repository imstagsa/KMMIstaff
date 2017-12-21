import os
from xlutils.filter import GlobReader,BaseFilter,DirectoryWriter,process

myfile='04-03_Replication_failure_control_doc.xls'
mydir='C://Python27//app' 

class MyFilter(BaseFilter): 

    goodlist = None
    
    def __init__(self,elist): 
        self.goodlist = goodlist
        self.wtw = 0
        self.wtc = 0
         
    def workbook(self, rdbook, wtbook_name): 
        self.next.workbook(rdbook, 'filtered_'+wtbook_name) 

    def row(self, rdrowx, wtrowx):
		pass

    def cell(self, rdrowx, rdcolx, wtrowx, wtcolx):
        value = self.rdsheet.cell(rdrowx,rdcolx).value
        if rdrowx == 1 and value != "DEL":
            if value in self.goodlist:
            self.wtc=self.wtc+1 
            self.next.row(rdrowx,wtrowx)
        else:
            return
        self.next.cell(rdrowx,rdcolx,self.wtc,wtcolx)
        
        
data = "DEL"
goodlist = data.split("\n")
process(GlobReader(os.path.join(mydir,myfile)),MyFilter(goodlist),DirectoryWriter(mydir))