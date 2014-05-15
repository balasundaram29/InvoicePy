
from  PyQt4 import QtGui,QtSql

class SQLComboBox(QtGui.QComboBox):
    def __init__(self,tableName,parent=None):
        super(SQLComboBox,self).__init__(parent)
        query=QtSql.QSqlQuery()
        s="SELECT * FROM `"+tableName+"`"
        query.exec_(s)
        nList=[]
        while(query.next()):
            v=query.value(1).toString()
            print v
            nList.append(v.__str__())
        self.insertItems(0,nList)
        for i in nList:
            print "buyer is",i
