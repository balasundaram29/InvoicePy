
#include <QSqlQuery>

#include "InvoiceCollection.h"
#include "QString"
#include "QVariant"
#include "QDebug"
#include "QSqlError"

from PyQt4 import  QtSql

def addBuyer(name,address,tinNo, cstNo, email):
    query=QtSql.QSqlQuery()
    query.prepare("REPLACE INTO Buyers(name,address,TINNo,CSTNo,email) VALUES(:name,:address,:TINNo,:CSTNo,:email)")
    query.bindValue(":name", name)
    query.bindValue(":address",address)
    query.bindValue(":TINNo",tinNo)
    query.bindValue(":CSTNo",cstNo)
    query.bindValue(":email",email)
    query.exec_()
    print query.lastError()


def addProduct( name,  unit,  type):
    query=QtSql.QSqlQuery()
    query.prepare("REPLACE INTO Products(name,unit,type) VALUES(:name,:unit,:type)")
    query.bindValue(":name", name)
    query.bindValue(":unit",unit)
    query.bindValue(":type",type)
    query.exec_()

def addInvoice(invno, invdate,  buyerID,cstrate,  scrate,  other):
    query=QtSql.QSqlQuery()
    query.prepare("REPLACE INTO Invoices(invno,invdate,buyerID,cstrate,scrate,other) VALUES(:invno,:invdate,:buyerID,:cstrate,:scrate,:other")
    #//"ON DUPLICATE KEY UPDATE"
   #//"invno=:invno,invdate=:invdate,buyerID=:buyerID,cstrate=:cstrate,scrate=:scrate,other=:other"
    #)
    query.bindValue(":invno", invno)
    query.bindValue(":invdate",invdate)
    query.bindValue(":buyerID",buyerID)
    query.bindValue(":cstrate",cstrate)
    query.bindValue(":scrate",scrate)
    query.bindValue(":other",other)
    
   
    query.exec_()
