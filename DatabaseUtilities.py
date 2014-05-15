
from PyQt4 import QtSql,QtGui,QtCore
def createConnection():
    db = QtSql.QSqlDatabase.addDatabase("QMYSQL")
    db.setHostName("localhost")
    db.setUserName("root")
    #/db.setPassword("root")
    if (not db.open()):
        QtGui.QMessageBox.critical(None,QtGui.qApp.tr("Database Error"), db.lastError().text())
        return False
    return True


def createTables():
    query=QtSql.QSqlQuery()
    #query.exec_("DROP DATABASE  IF EXISTS InvoiceAppDB")
    print '30'
    print query.lastError().text()

    print '31'
    print query.record().value(0).toInt()
    query.exec_("CREATE DATABASE IF NOT EXISTS InvoiceAppDB")
    print query.lastError()
    query.exec_("USE InvoiceAppDB")
    #query.exec_("DROP TABLE  IF EXISTS InvoicesAndProducts")
    #query.exec_("DROP TABLE  IF EXISTS Invoices")
    #query.exec_("DROP TABLE  IF EXISTS Products")
    #query.exec_("DROP TABLE  IF EXISTS Buyers")
    query.exec_(
            "CREATE TABLE IF NOT EXISTS Buyers "
            "("
            "buyerID INTEGER PRIMARY KEY AUTO_INCREMENT,"
            "name VARCHAR(40) NOT NULL UNIQUE,"
            "address VARCHAR(150),"
            "TINNo VARCHAR(25) ,"
            "CSTNo VARCHAR(25) , "
            "email VARCHAR(60)"
            ")ENGINE=InnoDB"
            )
    print query.lastError()
    #query.exec_("DROP TABLE  IF EXISTS Products")
    query.exec_(
            "CREATE TABLE IF NOT EXISTS Products"
            "("
            "productID INTEGER PRIMARY KEY AUTO_INCREMENT,"
            "name VARCHAR(40) NOT NULL UNIQUE,"
            "unit VARCHAR(20),"
            "type VARCHAR(15)"
            ")ENGINE=InnoDB"
            )
    print query.lastError()
    #query.exec_("INSERT INTO artist VALUES(1,'Ilayaraja','India'")
    #query.exec_("DROP TABLE  IF EXISTS Invoices")
    query.exec_(
            "CREATE TABLE IF NOT EXISTS Invoices"
            "("
            "invno INTEGER PRIMARY KEY,"
            "invdate DATE NOT NULL,"
            "buyerID INTEGER, "
            "cstrate DECIMAL(5,2),"
            "scrate DECIMAL(5,2),"
        #    "other INTEGER,"
            "bill_value INTEGER,"
            "tax_type  VARCHAR(6),"
            "form_c CHAR(1),"
            "FOREIGN KEY(buyerID) REFERENCES Buyers(buyerID)"
            ")ENGINE=InnoDB"
            )
    print query.lastError().text()
    #query.exec_("DROP TABLE  IF EXISTS InvoicesAndProducts")
    query.exec_(
            "CREATE TABLE IF NOT EXISTS InvoicesAndProducts"
            "("
            "id INTEGER PRIMARY KEY AUTO_INCREMENT,"
            "invno INTEGER NOT NULL,"
            "productID INTEGER NOT NULL,"
            "quantity INTEGER NOT NULL,"
            "price DECIMAL(10,2),"
            "FOREIGN KEY(invno) REFERENCES Invoices(invno),"
            "FOREIGN KEY(productID) REFERENCES Products(productID)"
            ")ENGINE=InnoDB"
            )
    print query.lastError()
    '''query.exec_(
            "REPLACE INTO InvoicesAndProducts"
            "("
            "id,"
            "invno,"
            "productID,"
            "quantity,"
            "price"
            ") VALUES"
            "(1,1,1,10,1000)"
            )*/
    print query.lastError()
    
}'''

def  setInnoDBEngine():
    query=QtSql.QSqlQuery()
    query.exec_(
            "ALTER TABLE  Buyers TYPE=InnoDB"
            )
    print query.lastError()
    query.exec_(
            "ALTER TABLE  Products TYPE=InnoDB"
            )
    
    print query.lastError()
    query.exec_(
            "ALTER TABLE  Invoices TYPE=InnoDB"
            )
    print query.lastError() 
    query.exec_(
            "ALTER TABLE  InvoicesAndProducts TYPE=InnoDB"
            )
    print query.lastError()
