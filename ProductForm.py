




from PyQt4 import QtGui,QtSql,QtCore
prID,prName,prUnit,prType=range(4)
class ViewWidget(QtGui.QWidget):
    def __init__(self,parent=None):
        super(ViewWidget,self).__init__(parent)
        layout=QtGui.QVBoxLayout()
        self.model=QtSql.QSqlTableModel()
        self.model.setTable('Products')

        self.table=QtGui.QTableView()
        self.table.setModel(self.model)
        self.model.select()
        self.table.setColumnHidden(prID,True)
        self.model.setHeaderData(prName,QtCore.Qt.Horizontal,"Name")
        self.model.setHeaderData(prUnit,QtCore.Qt.Horizontal,"Unit")
        self.model.setHeaderData(prType,QtCore.Qt.Horizontal,"Type")
        self.table.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)

        #table.setRowCount(10)
        #table.setColumnCount(4)
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
        label=QtGui.QLabel("Unit:")
        self.unitEdit=QtGui.QLineEdit()
        grid.addWidget(label,1,0)
        grid.addWidget(self.unitEdit,1,1)
        label=QtGui.QLabel("Type. :")
        self.typeEdit=QtGui.QLineEdit()
        grid.addWidget(label,2,0)
        grid.addWidget(self.typeEdit,2,1)

        lower=QtGui.QHBoxLayout()
        self.saveButton=QtGui.QPushButton("Save/Update")
        lower.addWidget(self.saveButton)
        self.deleteButton=QtGui.QPushButton("Delete")
        lower.addWidget(self.deleteButton)
        self.closeButton=QtGui.QPushButton("Close")
        lower.addWidget(self.closeButton)
        layout.addLayout(lower)
        self.setLayout(layout)

class ProductForm(QtGui.QDialog):
    def __init__(self,parent=None):
       super(ProductForm,self).__init__(parent)
       self.setWindowTitle("Manage Products")
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
       s='DELETE  FROM `Products` WHERE `productID` = '+"\'"+str(self.savedID)+"\'"
       q.exec_(s)


    def goForEdit(self):
       self.editing=True
       self.bookWidget.setCurrentIndex(1)

    def addButtonOfViewClicked(self):
       self.editing=False
       self.bookWidget.setCurrentIndex(1)

    def editButtonOfViewClicked(self):
       self.editing=True
       self.bookWidget.setCurrentIndex(1)
       id=self.viewWidget.model.record().value(1).toInt()[0]
       index=self.viewWidget.table.currentIndex()
       #index1=QtCore.QModelIndex()
       row=index.row()
       index=self.viewWidget.model.createIndex(row,1)
       dt=self.viewWidget.model.data(index).toString()
       s='SELECT * FROM `Products` WHERE `name` = '+ "\'" + dt +"\'"
       q=QtSql.QSqlQuery()
       q.exec_(s)

       if (q.next()):
           id=q.value(prID).toInt()[0]
           self.savedID=id
           v=q.value(prName).toString()
           self.editWidget.nameEdit.setText(v)
           v=q.value(prUnit).toString()
           self.editWidget.unitEdit.setText(v)
           v=q.value(prType).toString()
           self.editWidget.typeEdit.setText(v)


       edits=self.editWidget.findChildren(QtGui.QLineEdit)
       l=edits.__len__()
       print l

    def saveButtonOfEditClicked(self):
       q=QtSql.QSqlQuery()
       try:
           if (self.editing):
               print 'saved id is',self.savedID
               q.exec_('SET AUTOCOMMIT=0')
               q.exec_('START TRANSACTION')
               q.prepare(
                     'UPDATE `Products`'
                    '  SET `name`=:name,`unit`=:unit,`type`=:type '
                    'WHERE `productID`=:productID'
                    )
               q.bindValue(':productID',self.savedID)
               q.bindValue(':name',self.editWidget.nameEdit.text())
               q.bindValue(':unit',self.editWidget.unitEdit.text())
               q.bindValue(':type',self.editWidget.typeEdit.text())
               q.exec_()
               q.exec_("COMMIT")


           else:
               id=0
               q.exec_('SELECT MAX(`productID`) FROM `Products`')
               if (q.next()):
                    id=q.value(0).toInt()[0]+1
               q.prepare(
                     'INSERT  INTO `Products`'
                    ' (`productID`,`name`,`unit`,`type`)  '
                    'VALUES'
                    ' (:productID,:name,:unit,:type)  '
                    )
               q.bindValue(':productID:',id)
               q.bindValue(':name',self.editWidget.nameEdit.text())
               q.bindValue(':unit',self.editWidget.unitEdit.text())
               q.bindValue(':type',self.editWidget.typeEdit.text())
               q.exec_()
       except :pass


if  __name__=="__main__":

    app=QtGui.QApplication([])
    w=ProductForm()

    w.show()
    app.exec_()