
from PyQt4  import  QtGui,QtSql,QtCore

from InvoiceForm import MainWindow

import DatabaseUtilities as dbutils



def main():
    app=QtGui.QApplication([])
    dbutils.createConnection()
    dbutils.createTables()
    window=MainWindow(None)
    window.setGeometry(100,100,600,400)
    window.show()
    return app.exec_()

if __name__=='__main__':    
    main()
        