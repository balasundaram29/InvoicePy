




from PyQt4 import QtGui,QtSql,QtCore
bID,bName,bAddress,bTIN,bCST,bEmail=range(6)
class ViewWidget(QtGui.QWidget):
    def __init__(self,parent=None):
        super(ViewWidget,self).__init__(parent)
        layout=QtGui.QVBoxLayout()
        self.model=QtSql.QSqlTableModel()
        self.model.setTable('Buyers')
        self.table=QtGui.QTableView()
        self.table.setModel(self.model)
        self.model.select()
        self.table.setColumnHidden(bID,True)
        self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        layout.addWidget(self.table)
        lower=QtGui.QHBoxLayout()
        self.addButton=QtGui.QPushButton("Add")
        lower.addWidget(self.addButton)
        self.editButton=QtGui.QPushButton("Edit")
        lower.addWidget(self.editButton)
        self.deleteButton=QtGui.QPushButton("Delete")
        lower.addWidget(self.deleteButton)
        self.closeButton=QtGui.QPushButton("Close")
        lower.addWidget(self.closeButton)

        layout.addLayout(lower)
        self.setLayout(layout)
        id=self.model.record().value(1).toString()
        print "in View Widget", id
        index=self.table.currentIndex()
        dt=self.model.data(index).toInt()[0]
        print dt


class EditWidget(QtGui.QWidget):

    def  __init__(self,parent=None):
        super(EditWidget,self).__init__(parent)
        layout=QtGui.QVBoxLayout()
        grid=QtGui.QGridLayout()
        layout.addLayout(grid)
        label=QtGui.QLabel("Name:")
        self.nameEdit=QtGui.QLineEdit()
        grid.addWidget(label,0,0)
        grid.addWidget(self.nameEdit,0,1)
        label=QtGui.QLabel("Address:")
        self.addressEdit=QtGui.QTextEdit()
        grid.addWidget(label,1,0)
        grid.addWidget(self.addressEdit,1,1,1,1)
        label=QtGui.QLabel("TIN No. :")
        self.TINEdit=QtGui.QLineEdit()
        grid.addWidget(label,2,0)
        grid.addWidget(self.TINEdit,2,1)
        label=QtGui.QLabel("CST No. :")
        self.CSTEdit=QtGui.QLineEdit()
        grid.addWidget(label,3,0)
        grid.addWidget(self.CSTEdit,3,1)
        label=QtGui.QLabel("Email :")
        self.emailEdit=QtGui.QLineEdit()
        grid.addWidget(label,4,0)
        grid.addWidget(self.emailEdit,4,1)
        lower=QtGui.QHBoxLayout()
        self.saveButton=QtGui.QPushButton("Save/Update")
        lower.addWidget(self.saveButton)
        self.deleteButton=QtGui.QPushButton("Delete")
        lower.addWidget(self.deleteButton)
        self.closeButton=QtGui.QPushButton("Close")
        lower.addWidget(self.closeButton)
        layout.addLayout(lower)
        self.setLayout(layout)

class SQLTableForm(QtGui.QDialog):
    def __init__(self,parent=None):
       super(SQLTableForm,self).__init__(parent)
       self.setWindowTitle("Manage Buyers")
       self.editing=True
       self.savedID=None
       self.viewWidget=ViewWidget()
       self.editWidget=EditWidget()
       self.bookWidget=QtGui.QStackedWidget()
       self.bookWidget.addWidget(self.viewWidget)
       self.bookWidget.addWidget(self.editWidget)
       self.layout=QtGui.QVBoxLayout()
       self.layout.addWidget(self.bookWidget)
       self.setLayout(self.layout)
       self.connect(self.viewWidget.editButton,QtCore.SIGNAL('clicked()'),self.editButtonOfViewClicked)
       self.connect(self.viewWidget.deleteButton,QtCore.SIGNAL('clicked()'),self.editButtonOfViewClicked)
       self.connect(self.viewWidget.addButton,QtCore.SIGNAL('clicked()'),self.addButtonOfViewClicked)
       self.connect(self.editWidget.saveButton,QtCore.SIGNAL('clicked()'),self.saveButtonOfEditClicked)
       self.connect(self.editWidget.closeButton,QtCore.SIGNAL('clicked()'),lambda:self.bookWidget.setCurrentIndex(0))
       self.connect(self.viewWidget.table,QtCore.SIGNAL('doubleClicked(QModelIndex)'),self.editButtonOfViewClicked)
       self.connect(self.viewWidget.closeButton,QtCore.SIGNAL('clicked()'),self.close)

       self.connect(self.editWidget.deleteButton,QtCore.SIGNAL('clicked()'),self.deletButtonOfEditClicked)
    def deletButtonOfEditClicked(self):
       q=QtSql.QSqlQuery()
       s='DELETE  FROM `Buyers` WHERE `buyerID` = '+"\'"+str(self.savedID)+"\'"
       print s
       q.exec_(s)


    def goForEdit(self):
       print "double"
       self.editing=True
       self.bookWidget.setCurrentIndex(1)

    def addButtonOfViewClicked(self):
       self.editing=False
       self.bookWidget.setCurrentIndex(1)

    def editButtonOfViewClicked(self):
       self.editing=True
       self.bookWidget.setCurrentIndex(1)
       id=self.viewWidget.model.record().value(1).toInt()
       print id
       index=self.viewWidget.table.currentIndex()

       row=index.row()
       index=self.viewWidget.model.createIndex(row,1)
       dt=self.viewWidget.model.data(index).toString()
       print dt
       s='SELECT * FROM `Buyers` WHERE `name` = '+ "\'" + dt +"\'"
       q=QtSql.QSqlQuery()
       q.exec_(s)

       if (q.next()):
           id=q.value(bID).toInt()[0]
           self.savedID=id
           v=q.value(bName).toString()
           self.editWidget.nameEdit.setText(v)
           v=q.value(bAddress).toString()
           self.editWidget.addressEdit.setText(v)
           v=q.value(bCST).toString()
           self.editWidget.CSTEdit.setText(v)
           v=q.value(bTIN).toString()
           self.editWidget.TINEdit.setText(v)
           v=q.value(bEmail).toString()
           self.editWidget.emailEdit.setText(v)


       edits=self.editWidget.findChildren(QtGui.QLineEdit)
       l=edits.__len__()
       print l

    def saveButtonOfEditClicked(self):
       q=QtSql.QSqlQuery()
       print "in save button of edit of sqlTableForm"
       try:
           if (self.editing):
               print 'saved id is',self.savedID
               print self.editing
               v=self.editWidget.addressEdit.toPlainText()
               print "length of ",v,"is",str(v).__len__()
               q.exec_('SET AUTOCOMMIT=0')
               print q.lastError().text()
               q.exec_('START TRANSACTION')
               print q.lastError().text()

               q.prepare(
                     'UPDATE `Buyers`'
                    '  SET `name`=:name,`address`=:address,`TINNo`=:tin,`CSTNo`=:cst,`email`=:email  '
                    'WHERE `buyerID`=:buyerID'
                    )
               q.bindValue(':buyerID',self.savedID)
               q.bindValue(':name',self.editWidget.nameEdit.text())
               q.bindValue(':address',self.editWidget.addressEdit.toPlainText())
               q.bindValue(':tin',self.editWidget.TINEdit.text())
               q.bindValue(':cst',self.editWidget.CSTEdit.text())
               q.bindValue(':email',self.editWidget.emailEdit.text())
               q.exec_()
               print q.lastError().text()

               q.exec_("COMMIT")



           else:
               id=0
               q.exec_('SELECT MAX(`buyerID`) FROM `Buyers`')
               if (q.next()):
                    id=q.value(0).toInt()[0]+1

               q.prepare(
                     'INSERT  INTO `Buyers`'
                    ' (`buyerID`,`name`,`address`,`TINNo`,`CSTNo`,`email`)  '
                    'VALUES'
                    ' (:buyerID,:name,:address,:tin,:cst,:email)  '
                    )
               q.bindValue(':buyerID:',id)
               q.bindValue(':name',self.editWidget.nameEdit.text())
               q.bindValue(':address',self.editWidget.addressEdit.toPlainText())
               q.bindValue(':tin',self.editWidget.TINEdit.text())
               q.bindValue(':cst',self.editWidget.CSTEdit.text())
               q.bindValue(':email',self.editWidget.emailEdit.text())
               q.exec_()
       except :
           print "sql exception"
           print  q.lastError().text()


if  __name__=="__main__":

    app=QtGui.QApplication([])
    w=SQLTableForm()

    w.show()
    app.exec_()