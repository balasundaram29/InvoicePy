
from PyQt4 import QtGui,QtCore,QtSql
from InvoiceTableWidget import InvoiceTableWidget
from SQLComboBox import SQLComboBox
DESCPN=0
QTY=1
RATE=2
AMOUNT=3
invID=0
class ViewWidget(QtGui.QWidget):
    def __init__(self,parent=None):
        super(ViewWidget,self).__init__(parent)
        layout=QtGui.QVBoxLayout()
        self.model=QtSql.QSqlTableModel()
        self.model.setTable('Invoices')
        self.table=QtGui.QTableView()
        self.table.setModel(self.model)
        self.model.select()
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
        index=self.table.currentIndex()
        dt=self.model.data(index).toInt()[0]

class  EditWidget(QtGui.QWidget,QtCore.QObject):

    def __init__(self,parent=None):
        super(EditWidget,self).__init__(parent)
        self.INV_NO_KEY=0
        self.INV_DATE_KEY=1
        self.BUYER_ID_KEY=2
        self.TAX_RATE_KEY=3
        self.ROW_LIST_KEY=4
        self.SUM_KEY=5
        self.TAX_AMOUNT_KEY=6
        self.GRAND_TOTAL_KEY=7

        self.tableRowCount=0
        self.widgetValues={}
        self.productDict={}
        self.buyerDict={}
        grid=QtGui.QGridLayout()
        label=QtGui.QLabel('Invoice No :')
        self.invEdit=QtGui.QLineEdit()
        grid.addWidget(label,0,0)
        grid.addWidget(self.invEdit,0,1)
        label=QtGui.QLabel("Date:")
        self.dateEdit=QtGui.QDateEdit()
        self.dateEdit.setCalendarPopup(True)
        self.dateEdit.showMinimized()
        inDate=self.dateEdit.date()
        grid.addWidget(label,0,3)
        grid.addWidget(self.dateEdit,0,4)
        label=QtGui.QLabel("Buyer:")
        self.buyerCombo=SQLComboBox('Buyers')
        grid.addWidget(label,1,0)
        grid.addWidget(self.buyerCombo,1,1)
        label=QtGui.QLabel("Buyer's Address:")
        self.buyerAddressEdit=QtGui.QTextEdit()
        self.buyerAddressEdit.setMaximumHeight(60)
        grid.addWidget(label,1,3)
        grid.addWidget(self.buyerAddressEdit,1,4,1,2)
        self.itemTable=InvoiceTableWidget()

        self.addButton=QtGui.QPushButton("Add Item")
        grid.addWidget(self.addButton,2,4,1,1)
        self.delItemButton=QtGui.QPushButton("Delete Item")
        grid.addWidget(self.delItemButton,2,0,1,1)

        headers=['Item Description','Quantity','Rate','Amount']
        self.itemTable.setRowCount(0)
        self.itemTable.setColumnCount(4)
        self.itemTable.setHorizontalHeaderLabels(headers)
        self.itemTable.setSizePolicy(QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Expanding)
        self.itemTable.setAlternatingRowColors(True)

        header=self.itemTable.horizontalHeader()
        header.setResizeMode(QtGui.QHeaderView.Interactive)
        header.setResizeMode(0,QtGui.QHeaderView.Stretch)

        grid.addWidget(self.itemTable,3,0,1,5)
        label=QtGui.QLabel("CST/VAT:")
        self.taxEdit=QtGui.QLineEdit()
        self.taxEdit.setValidator(QtGui.QDoubleValidator(0,100,2,self))
        self.connect(self.taxEdit,QtCore.SIGNAL('textEdited(QString)'),self.computeTotal)

        grid.addWidget(label,4,0)
        grid.addWidget(self.taxEdit,4,1)
        label=QtGui.QLabel("Tax Type:")
        self.taxTypeCombo=QtGui.QComboBox()
        self.taxTypeCombo.insertItems(0,['CST','VAT'])
        grid.addWidget(label,4,2)
        grid.addWidget(self.taxTypeCombo,4,3)

        label=QtGui.QLabel("            Surcharge:")
        self.surchargeEdit=QtGui.QLineEdit()
        self.surchargeEdit.setValidator(QtGui.QDoubleValidator(0,100,2,self))
        self.connect(self.surchargeEdit,QtCore.SIGNAL('textEdited(QString)'),self.computeTotal)
        label.setBuddy(self.surchargeEdit)
        self.surchargeEdit.setText('0')
        self.formcCheckBox=QtGui.QCheckBox("Against Form 'C'")
        grid.addWidget(self.formcCheckBox,5,0)

        label=QtGui.QLabel('Total :')
        self.totalEdit=QtGui.QLineEdit()
        grid.addWidget(label,5,3)
        grid.addWidget(self.totalEdit,5,4)

        self.viewAllInvButton=QtGui.QPushButton('View All Invoices')
        self.printXlButton=QtGui.QPushButton('Print in Excel')
        self.printInFormButton=QtGui.QPushButton('Print in Form ')
        self.saveButton=QtGui.QPushButton('Save/Update')
        self.deleteButton=QtGui.QPushButton('Delete')
        grid.addWidget(self.viewAllInvButton,6,0)
        grid.addWidget(self.printXlButton,6,1)
        grid.addWidget(self.printInFormButton,6,2)
        grid.addWidget(self.saveButton,6,4)
        grid.addWidget(self.deleteButton,6,3)
        self.setLayout(grid)
        self.connect(self.addButton,QtCore.SIGNAL("clicked()"),self.addTableRow)
        self.connect(self.buyerCombo,QtCore.SIGNAL("currentIndexChanged(QString)"),self.setBuyerAddress)
        self.connect(self.itemTable,QtCore.SIGNAL("cellChanged(int,int)"),self.computeTotal)
        self.connect(self.saveButton,QtCore.SIGNAL("clicked()"),self.saveInvoice)
        self.connect(self.printXlButton,QtCore.SIGNAL("clicked()"),self.printInvoiceInExcel)
        self.connect(self.printInFormButton,QtCore.SIGNAL("clicked()"),self.printInvoiceInForm)
        self.connect(self.invEdit,QtCore.SIGNAL("returnPressed()"),self.loadInvoiceFromDB)
        self.connect(self.delItemButton,QtCore.SIGNAL("clicked()"),self.deleteTableRow)
        self.setBuyerAddress()

    def setBuyerAddress(self):
        buyerName=self.buyerCombo.currentText()
        query=QtSql.QSqlQuery()
        s="SELECT `address` from `Buyers` WHERE `name`="+"\'"+str(buyerName)+"\'"
        query.exec_(s)
        if query.next():
            buyadd=query.value(0).toString()
            self.buyerAddressEdit.setPlainText(buyadd)

    def deleteTableRow(self):
        row=self.itemTable.removeRow(self.itemTable.currentRow())
        self.computeTotal()

    def addTableRow(self):

        self.itemTable.insertRow(self.itemTable.rowCount())
        item=QtGui.QTableWidgetItem(self.tr('%1').arg(0.0))
        self.itemTable.setItem(self.itemTable.rowCount()-1,QTY,item)
        item=QtGui.QTableWidgetItem(self.tr('%1').arg(0.0))
        self.itemTable.setItem(self.itemTable.rowCount()-1,RATE,item)
        item=QtGui.QTableWidgetItem(self.tr('%1').arg(0.0))
        self.itemTable.setItem(self.itemTable.rowCount()-1,AMOUNT,item)
        c=SQLComboBox('Products')
        self.itemTable.setCellWidget(self.itemTable.rowCount()-1,DESCPN,c)
        self.computeTotal()

    def computeTotal(self):
        rowList=[]
        total=0.0
        grandTotal=0.0
        row=0
        for row in range(self.itemTable.rowCount()):
            if (self.itemTable.cellWidget(row,DESCPN)==None):
                return
            des=self.itemTable.cellWidget(row,DESCPN).currentText()
            qty=self.itemTable.item(row,QTY).text()
            rate=self.itemTable.item(row,RATE).text()
            singleRow=(des,qty,rate)
            rowList.append(singleRow)
            amount=qty.toFloat()[0]*rate.toFloat()[0]
            total=total+amount
            amtEdit=self.itemTable.item(row,AMOUNT)
            amtEdit.setData(QtCore.Qt.DisplayRole,QtCore.QString.number(amount,'f',2))
            singleRow=(des,qty,rate)
            amount=0.0
            rowList.append(singleRow)
            row=row+1

        tax=self.taxEdit.text().toDouble()[0]
        taxAmount=total*tax/100.00
        grandTotal=total+taxAmount

        gt=grandTotal
        self.totalEdit.setText(QtCore.QString.number(grandTotal,'f',2))
    def fillUpDictionaries(self):
        query=QtSql.QSqlQuery()
        s="SELECT `productID`,`name` from `Products`"
        query.exec_(s)
        while query.next():
            prID=query.value(0).toInt()[0]
            name=str(query.value(1).toString())
            self.productDict[name]=prID
        s="SELECT `buyerID`,`name` from `Buyers`"
        query.exec_(s)
        while query.next():
            buyID=query.value(0).toInt()[0]
            name=str(query.value(1).toString())
            self.buyerDict[name]=buyID
        invNo=(self.invEdit.text()).toInt()[0]
        self.widgetValues[self.INV_NO_KEY]=invNo
        invDate=self.dateEdit.date()
        self.widgetValues[self.INV_DATE_KEY]=invDate
        buyerName=self.buyerCombo.currentText()
        taxrate=self.taxEdit.text().toFloat()[0]
        self.widgetValues[self.TAX_RATE_KEY]=taxrate
        rowList=[]
        self.widgetValues[self.ROW_LIST_KEY]=rowList
        sum=0.0
        for row in range(self.itemTable.rowCount()):
            des=self.itemTable.cellWidget(row,DESCPN).currentText()
            rate=self.itemTable.item(row,RATE).text()
            qty=self.itemTable.item(row,QTY).text()
            amount=qty.toFloat()[0]*rate.toFloat()[0]
            sum=sum+amount
            singleRow=(des,qty,rate,amount)
            rowList.append(singleRow)
        self.widgetValues[self.SUM_KEY]=sum
        self.widgetValues[self.TAX_AMOUNT_KEY]=sum*self.widgetValues[self.TAX_RATE_KEY]/100.0
        self.widgetValues[self.GRAND_TOTAL_KEY]=self.widgetValues[self.SUM_KEY]+self.widgetValues[self.TAX_AMOUNT_KEY]

    def saveInvoice(self):
        if (self.itemTable.rowCount()==0):
            return
        self.fillUpDictionaries()
        query=QtSql.QSqlQuery()
        s="SELECT * FROM `Invoices` WHERE `invno` = "+ "\'" +str(self.widgetValues[self.INV_NO_KEY]) +"\'"
        query.exec_(s)
        if(query.next()):#generate inv no exists .overwrite? prompt
            pass
        query.exec_('SET AUTOCOMMIT=0')
        query.exec_('START TRANSACTION')
        s="DELETE   FROM `InvoicesAndProducts` WHERE " \
          "`invno` = " + "\'"+str(self.widgetValues[self.INV_NO_KEY])+"\'"
        query.exec_(s)
        s="DELETE   FROM `Invoices` WHERE " \
          "`invno` = " + "\'"+str(self.widgetValues[self.INV_NO_KEY])+"\'"
        query.exec_(s)

        query.prepare(
             'INSERT INTO `Invoices`'
            ' (`invno`,`invdate`,`buyerID`,`cstrate`,`scrate`, `bill_value`,`tax_type`,`form_c`)  '
            'VALUES'
            ' (:invno,:invdate,:buyerID,:cstrate,:scrate,:bill_value,:tax_type,:form_c)  '
            )
        query.bindValue(':invno',self.widgetValues[self.INV_NO_KEY])
        query.bindValue(':invdate',self.widgetValues[self.INV_DATE_KEY])
        query.bindValue(':buyerID',self.buyerDict[str(self.buyerCombo.currentText())])
        query.bindValue(':cstrate',self.widgetValues[self.TAX_RATE_KEY])
        query.bindValue(':scrate',0)#not used now
        query.bindValue(':bill_value',self.widgetValues[self.GRAND_TOTAL_KEY])
        query.bindValue(':tax_type',str(self.taxTypeCombo.currentText()))
        formCTxt='N'
        if(self.formcCheckBox.isChecked()):
            formCTxt="Y"
        query.bindValue(':form_c',formCTxt)
        query.exec_()

        for row in self.widgetValues[self.ROW_LIST_KEY]:
            query.prepare(
             'INSERT INTO `InvoicesAndProducts`'
            ' (`invno`,`productID`,`quantity`,`price`)  '
            'VALUES'
            '(:ino,:prId,:qty,:price)'
            )
            query.bindValue(':ino',self.widgetValues[self.INV_NO_KEY])
            query.bindValue(':prID',self.productDict[str(row[DESCPN])])
            query.bindValue(':qty',row[QTY])
            query.bindValue((':price'),row[RATE])
            query.exec_()
        query.exec_('COMMIT')
    def printInvoiceInForm(self):
        print "Hello print in Form"
        printer=QtGui.QPrinter()
        printer.setPageMargins(0.0,0.0,0.0,0.0,QtGui.QPrinter.Millimeter)
        painter=QtGui.QPainter()
        points=printer.resolution()
        dialog=QtGui.QPrintDialog(printer)
        if(dialog.exec_()==QtGui.QDialog.Accepted):
            painter.begin(printer)
            INCHES_TO_MM=25.4
            INV_REF_X=10
            INV_REF_Y=20
            option=QtGui.QTextOption()
            option.setAlignment(QtCore.Qt.AlignTop|QtCore.Qt.AlignRight)
            option.setWrapMode((QtGui.QTextOption.WordWrap))
            rect=QtCore.QRectF(0,0,points*2.0,points*5.0)
            #painter.drawText(QtCore.QRectF(points*(INV_REF_X*INCHES_TO_MM,INV_REF_Y*INCHES_TO_MM),points*0.0,points*2.0,points/2.0),"hello World,printer")
            painter.drawRect(rect)
            import module1
            s=module1.changeAnywhere(1111)
            painter.drawText(rect,s+'top',option)
            option.setAlignment(QtCore.Qt.AlignBottom|QtCore.Qt.AlignLeft)
            option.setWrapMode((QtGui.QTextOption.WordWrap))
            painter.drawText(rect,s+'bottom',option)
            painter.end()

    def printInExcelAnnaiFormat(self):
        self.fillUpDictionaries()
        from win32com.client import Dispatch,constants
        import PageSetup
        import string
        XL_EDGE_TOP=10
        XL_CONT_LINE=1
        excelApp=Dispatch('Excel.Application')

        book=excelApp.Workbooks.Add()
        sheet1=book.Worksheets(1)
        sheet1.Name='Original'
        sheet2=book.Worksheets(2)
        sheet2.Name='Duplicate'
        sheet3=book.Worksheets(3)
        sheet3.Name='Triplicate'
        sheets=book.Worksheets
        sheet4=sheets.Add(None,sheet3)
        sheet4.Name="Quadruplicate"
        sheets=[]
        #sheets.append(sheet1)
        sheets=book.Worksheets
        copyTypeList=['Original','Duplicate','Triplicate','Quadruplicate']
        sheetIndex=-1
        excelApp.Visible=True
#        excelApp.Editable=False

        for sheet in sheets:
            sheetIndex=sheetIndex+1
            SNO=0
            START_ROW=1
            START_COL=1
            #we will be thinking in terms of current row and column
            CURR_ROW=1
            CURR_COL=1
            FINAL_ROW=53
            FINAL_COL=9
            INVNO_COL=7
            INV_DATE_COL=9
            SNO_COL=CURR_COL
            DESCPN_COL=SNO_COL+1
            QUANTITY_COL=DESCPN_COL+5
            RATE_COL=QUANTITY_COL+1
            AMOUNT_COL=RATE_COL+1
            PageSetup.setupPage(sheet,copyTypeList[sheetIndex],CURR_ROW,FINAL_COL)
            CURR_ROW=CURR_ROW+1


            pics=sheet.Pictures().Insert(r"C:\cris1.jpg")
            pics.Left=sheet.Cells(2,1).Left
            pics.Top=sheet.Cells(2,1).Top+20
            pics.ShapeRange.LockAspectRatio=True
            pics.ShapeRange.Width=100


            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
            xlRange.ColumnWidth=11
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
            xlRange.ColumnWidth=17
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.ColumnWidth=20

            #sheet.Cells(CURR_ROW,CURR_COL).Value="CHRISTECH ENGINEERING COMPANY"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=22
            xlRange.Value="CHRISTECH ENGINEERING COMPANY"


            #draw outer borders of page
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.VerticalAlignment=constants.xlTop


            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL),sheet.Cells(FINAL_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,FINAL_COL),sheet.Cells(FINAL_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,START_COL),sheet.Cells(CURR_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin




            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="   Mfrs.:Openwell,Borewell Submersible Pumpsets,Jet And Monoblock Pumpsets,"
            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="Single Phase And Three Phase Industrial Motors"

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="3/32A, J.V.Nagar,VK Road,Vinayagapuram,"
            sheet.Cells(CURR_ROW,FINAL_COL).Value="TIN : 33942227422 "


            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="Saravanampatti Post, Coimbatore - 641 035"
            sheet.Cells(CURR_ROW,FINAL_COL).Value="CST NO : 1164697 "

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="Mobile : 95669 95507"
            sheet.Cells(CURR_ROW,FINAL_COL).Value="Date : 13.11.2013"

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL+2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlCenter
           # xlRange.Value="E-mail: christec012@gmail.com"
            sheet.Cells(CURR_ROW,FINAL_COL).Value=""
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-3),sheet.Cells(CURR_ROW+5,FINAL_COL-3))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(START_ROW+1,START_ROW),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Font.Bold=True
            CURR_ROW=CURR_ROW+1

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=13
            xlRange.Value="To :"

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=22
            xlRange.Font.Bold=True
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.Value="INVOICE"
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin


            '''xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            '''
            buyerName=self.buyerCombo.currentText()
            query=QtSql.QSqlQuery()
            s="SELECT * from `Buyers` WHERE `name`="+"\'"+buyerName+"\'"
            query.exec_(s)
            if query.next():
                buyID=query.value(0).toInt()[0]
                buyAddress=str(query.value(2).toString())
                addressAsList=string.split(buyAddress,'\n')

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=14
            xlRange.Value=str(buyerName)

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=13
            if len(addressAsList)>= 1 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[0]

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,INVNO_COL),sheet.Cells(CURR_ROW,INVNO_COL+1))
            xlRange.Merge()
            xlRange.Font.Size=16

            sheet.Cells(CURR_ROW,INVNO_COL).Value="Invoice No :"+str(self.invEdit.text())
            sheet.Cells(CURR_ROW,INVNO_COL).Font.Bold=True
            sheet.Cells(CURR_ROW,INV_DATE_COL).Value="Date:"+str(self.dateEdit.date().toString('dd-MM-yyyy'))
            sheet.Cells(CURR_ROW,INV_DATE_COL).Font.Size=15
            sheet.Cells(CURR_ROW,INV_DATE_COL).Font.Bold=True
            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12
            if len(addressAsList)>= 2 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[1]

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12
            if len(addressAsList)>= 3 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[2]

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin


            CURR_ROW=CURR_ROW+1
            TINNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=12
            tinstring=str(query.value(3).toString())
            sheet.Cells(TINNO_ROW,CURR_COL).Value="Buyers TIN No :"+tinstring
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL+4),sheet.Cells(CURR_ROW,CURR_COL+5))
            xlRange.Merge()
            xlRange.Font.Size=12
            sheet.Cells(TINNO_ROW,CURR_COL+4).Value="CST :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="D.C.No : "
            sheet.Cells(CURR_ROW,FINAL_COL).Font.Size=12
            sheet.Cells(CURR_ROW,FINAL_COL).Value="Date: "
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            CURR_ROW=CURR_ROW+1
            TINNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=12
            tinstring=str(query.value(3).toString())
            sheet.Cells(TINNO_ROW,CURR_COL).Value="Transporter :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL+4),sheet.Cells(CURR_ROW,CURR_COL+5))
            xlRange.Merge()
            xlRange.Font.Size=12
            sheet.Cells(TINNO_ROW,CURR_COL+4).Value="LR No :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="Order No : "
            sheet.Cells(CURR_ROW,FINAL_COL).Font.Size=12
            sheet.Cells(CURR_ROW,FINAL_COL).Value="Date: "
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            CURR_ROW=CURR_ROW+1
            TINNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=12
            tinstring=str(query.value(3).toString())
            sheet.Cells(TINNO_ROW,CURR_COL).Value="Destination :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL+4),sheet.Cells(CURR_ROW,CURR_COL+5))
            xlRange.Merge()
            xlRange.Font.Size=12
            sheet.Cells(TINNO_ROW,CURR_COL+4).Value="Freight :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="Documents Through: "
            #sheet.Cells(CURR_ROW,FINAL_COL).Font.Size=12
            #sheet.Cells(CURR_ROW,FINAL_COL).Value="Date: "
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW-2,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL-1))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin







            CURR_ROW=CURR_ROW+1
            ITEM_TABLE_FINAL_ROW=47
            #draw item table borders
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))

            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin
            xlRange.Font.Bold=True
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(ITEM_TABLE_FINAL_ROW,FINAL_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(ITEM_TABLE_FINAL_ROW,START_COL),sheet.Cells(CURR_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(CURR_ROW,SNO_COL))
            xlRange.Font.Size=12
            xlRange.ColumnWidth=4
            sheet.Cells(CURR_ROW,SNO_COL).Value="SNo"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,SNO_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlLeft
            sheet.Cells(CURR_ROW,DESCPN_COL).Value="Description"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,QUANTITY_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlLeft
            sheet.Cells(CURR_ROW,QUANTITY_COL).Value="Quantity"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,QUANTITY_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlRight
            sheet.Cells(CURR_ROW,RATE_COL).Value="Rate"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,RATE_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlRight
            sheet.Cells(CURR_ROW,AMOUNT_COL).Value="Amount"
            rowList=[]



            ROWS_PER_ITEM=3
            CURR_ROW=CURR_ROW+1
            SAVED_ROW=CURR_ROW
            MAX_ITEMS=9
            sno=0

            rowList=self.widgetValues[self.ROW_LIST_KEY]
            for row in rowList:
                sno=sno+1
                xlRange=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW+ROWS_PER_ITEM-1,QUANTITY_COL-1))
                xlRange.Merge()
                xlRange.VerticalAlignment = constants.xlTop
                xlRange.WrapText = True
                sheet.Cells(CURR_ROW,SNO_COL).Value=sno
                sheet.Cells(CURR_ROW,SNO_COL).HorizontalAlignment=constants.xlCenter
                sheet.Cells(CURR_ROW,DESCPN_COL).Value=str(row[DESCPN])
                sheet.Cells(CURR_ROW,QUANTITY_COL).Value=str(row[QTY])
                sheet.Cells(CURR_ROW,QUANTITY_COL).HorizontalAlignment=constants.xlCenter
                #rate_prg=row[RATE].toFloat()[0]
                #num=QtCore.QString.number(rate_prg,'f',2)
                sheet.Cells(CURR_ROW,RATE_COL).Value=str(row[RATE])

                sheet.Cells(CURR_ROW,AMOUNT_COL).Value=str(row[AMOUNT])
                if int(float(str(row[RATE])) == 0):
                    sheet.Cells(CURR_ROW,RATE_COL).Value='---'
                    sheet.Cells(CURR_ROW,AMOUNT_COL).Value='---'
                CURR_ROW=CURR_ROW+ROWS_PER_ITEM

            TOTAL_AMOUNT_ROW=43
            sheet.Cells(TOTAL_AMOUNT_ROW,RATE_COL).Value="TOTAL"

            sheet.Cells(TOTAL_AMOUNT_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight

            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).Value=self.widgetValues[self.SUM_KEY]
            xlRange=sheet.Range(sheet.Cells(TOTAL_AMOUNT_ROW,RATE_COL),sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL))
            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin

            cst=self.taxEdit.text().toFloat()[0]
            TAX_ROW=TOTAL_AMOUNT_ROW+1
            if(self.formcCheckBox.isChecked()):
                xlRange=sheet.Range(sheet.Cells(TAX_ROW,DESCPN_COL),sheet.Cells(TAX_ROW,QUANTITY_COL-1))
                xlRange.Merge()
                xlRange.Value="Against  Form 'C'"
                xlRange.HorizontalAlignment=constants.xlRight

            sheet.Cells(TAX_ROW,RATE_COL).Value=str(self.taxTypeCombo.currentText())+' : '+str(cst)+'%'
            sheet.Cells(TAX_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TAX_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(TAX_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TAX_ROW,AMOUNT_COL).Value=self.widgetValues[self.TAX_AMOUNT_KEY]
            xlRange=sheet.Range(sheet.Cells(TAX_ROW,START_COL),sheet.Cells(TAX_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(TAX_ROW,RATE_COL),sheet.Cells(TAX_ROW,AMOUNT_COL))
            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin



            GRAND_TOTAL_ROW=TAX_ROW+1
            xlRange=sheet.Range(sheet.Cells(GRAND_TOTAL_ROW,START_COL),sheet.Cells(GRAND_TOTAL_ROW+2,QUANTITY_COL))
            xlRange.Merge()
            xlRange.VerticalAlignment=constants.xlCenter
            #xlRange.Value="Total Amount In Words : "
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            import module1,math,string
            s=module1.changeAnywhere(math.floor(self.widgetValues[self.GRAND_TOTAL_KEY]))
            s=string.replace(s,"Zero","")
            xlRange.WrapText = True
            xlRange.Value="Rupees "+s+" Only."
            xlRange.Font.Bold=True

            xlRange=sheet.Range(sheet.Cells(GRAND_TOTAL_ROW,RATE_COL),sheet.Cells(GRAND_TOTAL_ROW+2,RATE_COL))
            xlRange.Merge()
            xlRange.VerticalAlignment=constants.xlCenter
            xlRange.Value="GRAND TOTAL"
            xlRange.Font.Bold=True

            xlRange=sheet.Range(sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL),sheet.Cells(GRAND_TOTAL_ROW+2,AMOUNT_COL))
            xlRange.Merge()
            xlRange.Font.Bold=True
            xlRange.VerticalAlignment=constants.xlCenter
            sheet.Cells(GRAND_TOTAL_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).Value=self.widgetValues[self.GRAND_TOTAL_KEY]

            SIGN_FIRST_ROW=GRAND_TOTAL_ROW+3
            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW,QUANTITY_COL),sheet.Cells(SIGN_FIRST_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.Font.Bold=True
            xlRange.HorizontalAlignment = constants.xlRight
            xlRange.Value="for Christec Engineering Company"

            SIGN_FINAL_ROW=FINAL_ROW
            xlRange=sheet.Range(sheet.Cells(SIGN_FINAL_ROW,QUANTITY_COL),sheet.Cells(SIGN_FINAL_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.Font.Bold=True
            xlRange.HorizontalAlignment = constants.xlRight
            xlRange.Value="Authorized Signatory"
            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW,CURR_COL),sheet.Cells(SIGN_FIRST_ROW,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.HorizontalAlignment = constants.xlLeft
            xlRange.Value="Subject to Coimbatore jurisdiction.  "
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin


            import module1,math,string
            s=module1.changeAnywhere(math.floor(self.widgetValues[self.GRAND_TOTAL_KEY]))
            s=string.replace(s,"Zero","")

            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW+1,CURR_COL),sheet.Cells(SIGN_FIRST_ROW+2,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.VerticalAlignment = constants.xlTop
            xlRange.WrapText = True
            xlRange.Value="Interest at 24% per annum will be charged on amounts paid within 30 days from the date of invoice. "
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW-2,START_COL),sheet.Cells(FINAL_ROW,START_COL+2))
            xlRange.Merge()
            xlRange.VerticalAlignment = constants.xlTop
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Value="Prepared By : "
            xlRange=sheet.Range(sheet.Cells(FINAL_ROW-2,START_COL+3),sheet.Cells(FINAL_ROW,START_COL+5))
            xlRange.Merge()
            xlRange.VerticalAlignment = constants.xlTop
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Value="Checked By : "


        excelApp.Visible=True

    def printInvoiceInExcel(self):
        self.printInExcelAnnaiFormat()
        return
        self.fillUpDictionaries()
        from win32com.client import Dispatch,constants
        import PageSetup
        import string
        XL_EDGE_TOP=10
        XL_CONT_LINE=1
        excelApp=Dispatch('Excel.Application')

        book=excelApp.Workbooks.Add()
        sheet1=book.Worksheets(1)
        sheet1.Name='Original'
        sheet2=book.Worksheets(2)
        sheet2.Name='Duplicate'
        sheet3=book.Worksheets(3)
        sheet3.Name='Triplicate'
        sheets=book.Worksheets
        sheet4=sheets.Add(None,sheet3)
        sheet4.Name="Quadruplicate"
        sheets=book.Worksheets
        copyTypeList=['Original','Duplicate','Triplicate','Quadruplicate']
        sheetIndex=-1
        excelApp.Visible=True
        excelApp.Editable=False

        for sheet in sheets:
            sheetIndex=sheetIndex+1
            SNO=0
            START_ROW=1
            START_COL=1
            #we will be thinking in terms of current row and column
            CURR_ROW=1
            CURR_COL=1
            FINAL_ROW=52
            FINAL_COL=9
            INVNO_COL=7
            INV_DATE_COL=9
            SNO_COL=CURR_COL
            DESCPN_COL=SNO_COL+1
            QUANTITY_COL=DESCPN_COL+5
            RATE_COL=QUANTITY_COL+1
            AMOUNT_COL=RATE_COL+1
            PageSetup.setupPage(sheet,copyTypeList[sheetIndex],CURR_ROW,FINAL_COL)
            CURR_ROW=CURR_ROW+1

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
            xlRange.ColumnWidth=11
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
            xlRange.ColumnWidth=17
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.ColumnWidth=20

            sheet.Cells(CURR_ROW,CURR_COL).Value="INVOICE"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=22
            xlRange.Font.Bold=True
            xlRange.HorizontalAlignment=constants.xlCenter
            xlRange.VerticalAlignment=constants.xlTop

            #draw outer borders of page
            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(START_ROW,FINAL_COL),sheet.Cells(FINAL_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,FINAL_COL),sheet.Cells(FINAL_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,START_COL),sheet.Cells(START_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin




            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="Christec Engineering Company"


            xlRange=sheet.Range(sheet.Cells(CURR_ROW,INVNO_COL),sheet.Cells(CURR_ROW,INVNO_COL+1))
            xlRange.Merge()
            xlRange.Font.Size=16

            sheet.Cells(CURR_ROW,INVNO_COL).Value="Invoice No :"+str(self.invEdit.text())
            sheet.Cells(CURR_ROW,INV_DATE_COL).Value="Date:"+str(self.dateEdit.date().toString('dd-MM-yyyy'))
            sheet.Cells(CURR_ROW,INV_DATE_COL).Font.Size=15
            CURR_ROW=CURR_ROW+1


            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-4))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="3/32A,J.V. Nagar"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="Buyer's Order No: "

            CURR_ROW=CURR_ROW+1

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="VK Road,Vinayagapuram"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="Delivered Through :  "

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="Saravanampatti Post"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL-2),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.Value="Documents Through :  "
            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="Coimbatore-641 035"

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=16
            sheet.Cells(CURR_ROW,CURR_COL).Value="Mobile : 95669 95507"

            CURR_ROW=CURR_ROW+1
            TINNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=14
            sheet.Cells(TINNO_ROW,CURR_COL).Value="TIN No : 33942227422"
            CURR_ROW=CURR_ROW+1
            CSTNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=14
            sheet.Cells(CURR_ROW,CURR_COL).Value="CST NO : 1164697 dt 13.11.2013"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=14

            CURR_ROW=CURR_ROW+3
            sheet.Cells(CURR_ROW,CURR_COL).Value="Buyer's Name and Address :"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12
            buyerName=self.buyerCombo.currentText()
            query=QtSql.QSqlQuery()
            s="SELECT * from `Buyers` WHERE `name`="+"\'"+buyerName+"\'"
            query.exec_(s)
            if query.next():
                buyID=query.value(0).toInt()[0]
                buyAddress=str(query.value(2).toString())
                addressAsList=string.split(buyAddress,'\n')

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12

            sheet.Cells(CURR_ROW,CURR_COL).Value=str(buyerName)

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=13
            if len(addressAsList)>= 1 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[0]

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12
            if len(addressAsList)>= 2 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[1]

            CURR_ROW=CURR_ROW+1
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
            xlRange.Merge()
            xlRange.Font.Size=12
            if len(addressAsList)>= 3 :
                sheet.Cells(CURR_ROW,CURR_COL).Value=addressAsList[2]

            CURR_ROW=CURR_ROW+1
            TINNO_ROW=CURR_ROW
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
            xlRange.Merge()
            xlRange.Font.Size=12
            tinstring=str(query.value(3).toString())
            sheet.Cells(TINNO_ROW,CURR_COL).Value="Buyers TIN No :"+tinstring

            CURR_ROW=CURR_ROW+2
            ITEM_TABLE_FINAL_ROW=47
            #draw item table borders
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,START_COL),sheet.Cells(CURR_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeTop).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,FINAL_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,FINAL_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(ITEM_TABLE_FINAL_ROW,FINAL_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(ITEM_TABLE_FINAL_ROW,START_COL),sheet.Cells(CURR_ROW,START_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(CURR_ROW,SNO_COL))
            xlRange.Font.Size=12
            xlRange.ColumnWidth=4
            sheet.Cells(CURR_ROW,SNO_COL).Value="SNo"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,SNO_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlLeft
            sheet.Cells(CURR_ROW,DESCPN_COL).Value="Description"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,QUANTITY_COL))
            xlRange.Borders(constants.xlEdgeLeft).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeLeft).Weight=constants.xlThin
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlLeft
            sheet.Cells(CURR_ROW,QUANTITY_COL).Value="Quantity"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,QUANTITY_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlRight
            sheet.Cells(CURR_ROW,RATE_COL).Value="Rate"
            xlRange=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(ITEM_TABLE_FINAL_ROW,RATE_COL))
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.Font.Size=12
            xlRange.HorizontalAlignment=constants.xlRight
            sheet.Cells(CURR_ROW,AMOUNT_COL).Value="Amount"
            rowList=[]



            ROWS_PER_ITEM=3
            CURR_ROW=CURR_ROW+1
            SAVED_ROW=CURR_ROW
            MAX_ITEMS=9
            sno=0

            rowList=self.widgetValues[self.ROW_LIST_KEY]
            for row in rowList:
                sno=sno+1
                xlRange=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW+ROWS_PER_ITEM-1,QUANTITY_COL-1))
                xlRange.Merge()
                xlRange.VerticalAlignment = constants.xlTop
                xlRange.WrapText = True
                sheet.Cells(CURR_ROW,SNO_COL).Value=sno
                sheet.Cells(CURR_ROW,SNO_COL).HorizontalAlignment=constants.xlCenter
                sheet.Cells(CURR_ROW,DESCPN_COL).Value=str(row[DESCPN])
                sheet.Cells(CURR_ROW,QUANTITY_COL).Value=str(row[QTY])
                sheet.Cells(CURR_ROW,QUANTITY_COL).HorizontalAlignment=constants.xlCenter
                #rate_prg=row[RATE].toFloat()[0]
                #num=QtCore.QString.number(rate_prg,'f',2)
                sheet.Cells(CURR_ROW,RATE_COL).Value=str(row[RATE])

                sheet.Cells(CURR_ROW,AMOUNT_COL).Value=str(row[AMOUNT])
                if int(float(str(row[RATE])) == 0):
                    sheet.Cells(CURR_ROW,RATE_COL).Value='---'
                    sheet.Cells(CURR_ROW,AMOUNT_COL).Value='---'
                CURR_ROW=CURR_ROW+ROWS_PER_ITEM

            TOTAL_AMOUNT_ROW=45
            sheet.Cells(TOTAL_AMOUNT_ROW,RATE_COL).Value="TOTAL"
            sheet.Cells(TOTAL_AMOUNT_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight

            sheet.Cells(TOTAL_AMOUNT_ROW,AMOUNT_COL).Value=self.widgetValues[self.SUM_KEY]

            cst=self.taxEdit.text().toFloat()[0]

            TAX_ROW=TOTAL_AMOUNT_ROW+1
            if(self.formcCheckBox.isChecked()):
                xlRange=sheet.Range(sheet.Cells(TAX_ROW,DESCPN_COL),sheet.Cells(TAX_ROW,QUANTITY_COL-1))
                xlRange.Merge()
                xlRange.Value="Against  Form 'C'"
                xlRange.HorizontalAlignment=constants.xlRight

            sheet.Cells(TAX_ROW,RATE_COL).Value=str(self.taxTypeCombo.currentText())+' : '+str(cst)+'%'
            sheet.Cells(TAX_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TAX_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(TAX_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(TAX_ROW,AMOUNT_COL).Value=self.widgetValues[self.TAX_AMOUNT_KEY]


            GRAND_TOTAL_ROW=TAX_ROW+1
            sheet.Cells(GRAND_TOTAL_ROW,RATE_COL).Value="GRAND TOTAL"
            sheet.Cells(GRAND_TOTAL_ROW,RATE_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).NumberFormat="0.00"
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).HorizontalAlignment=constants.xlRight
            sheet.Cells(GRAND_TOTAL_ROW,AMOUNT_COL).Value=self.widgetValues[self.GRAND_TOTAL_KEY]

            SIGN_FIRST_ROW=GRAND_TOTAL_ROW+1
            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW,QUANTITY_COL),sheet.Cells(SIGN_FIRST_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.HorizontalAlignment = constants.xlRight
            xlRange.Value="for Christec Engineering Company"

            SIGN_FINAL_ROW=FINAL_ROW
            xlRange=sheet.Range(sheet.Cells(SIGN_FINAL_ROW,QUANTITY_COL),sheet.Cells(SIGN_FINAL_ROW,AMOUNT_COL))
            xlRange.Merge()
            xlRange.HorizontalAlignment = constants.xlRight
            xlRange.Value="Authorized Signatory"
            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW,CURR_COL),sheet.Cells(SIGN_FIRST_ROW,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.HorizontalAlignment = constants.xlLeft
            xlRange.Value="Total Amount In Words : "
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin


            import module1,math,string
            s=module1.changeAnywhere(math.floor(self.widgetValues[self.GRAND_TOTAL_KEY]))
            s=string.replace(s,"Zero","")

            xlRange=sheet.Range(sheet.Cells(SIGN_FIRST_ROW+1,CURR_COL),sheet.Cells(SIGN_FINAL_ROW,QUANTITY_COL-1))
            xlRange.Merge()
            xlRange.VerticalAlignment = constants.xlTop
            xlRange.WrapText = True
            xlRange.Value="Rupees "+s+" Only."
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Borders(constants.xlEdgeBottom).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeBottom).Weight=constants.xlThin

            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,START_COL),sheet.Cells(FINAL_ROW,START_COL+2))
            xlRange.Merge()
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Value="Prepared By : "
            xlRange=sheet.Range(sheet.Cells(FINAL_ROW,START_COL+3),sheet.Cells(FINAL_ROW,START_COL+5))
            xlRange.Merge()
            xlRange.Borders(constants.xlEdgeRight).LineStyle=constants.xlContinuous
            xlRange.Borders(constants.xlEdgeRight).Weight=constants.xlThin
            xlRange.Value="Checked By : "


        excelApp.Visible=True


    def clearAllNow(self):
        widlist=self.findChildren(QtGui.QLineEdit)
        for wid in widlist:
            wid.clear()
        rowcount=self.itemTable.rowCount()
        for row in range(rowcount):
            self.itemTable.removeRow(0)
        self.dateEdit.setDate(QtCore.QDate.currentDate())
        self.buyerCombo.setCurrentIndex(0)
        self.setBuyerAddress()

    def loadInvoiceFromDB(self):

        invNo=(self.invEdit.text()).toInt()[0]
        self.clearAllNow()
        self.invEdit.setText(str(invNo))
        query=QtSql.QSqlQuery()
        s="SELECT * FROM `Invoices` WHERE `invno` = "+ "\'" + str(invNo) +"\'"
        query.exec_(s)
        if(query.next()):
            self.dateEdit.setDate(query.value(1).toDate())
            buyerID=query.value(2).toInt()[0]
            self.taxEdit.setText(str(query.value(3).toFloat()[0]))
            self.surchargeEdit.setText(query.value(4).toString()[0])
            self.taxTypeCombo.setCurrentIndex(self.taxTypeCombo.findText(query.value(6).toString()))
            if (query.value(7).toString()=="Y"):
                self.formcCheckBox.setChecked(True)
            else:
                self.formcCheckBox.setChecked(False)

            query3=QtSql.QSqlQuery()
            s='SELECT `name` FROM `Buyers` WHERE `buyerID` = '+ "\'" + str(buyerID) +"\'"
            query3.exec_(s)
            if(query3.next()):
                buyerName=query3.value(0).toString()
                ix=self.buyerCombo.findText(buyerName)
                self.buyerCombo.setCurrentIndex(ix)
        else:
            self.clearAllNow()
            self.invEdit.setText(str(invNo))
        s="SELECT * FROM `InvoicesAndProducts` WHERE `invno` = "+ "\'" + str(invNo) +"\'"
        query.exec_(s)
        row=0
        while query.next():
            self.addTableRow()

            prID=query.value(2).toInt()[0]
            query2=QtSql.QSqlQuery()
            s="SELECT `name` from `Products` WHERE `productID` = "+"\'"+ str(prID) +"\'"
            query2.exec_(s)
            if query2.next():
                prName=query2.value(0).toString()
                combo=self.itemTable.cellWidget(row,DESCPN)
                ix=combo.findText(prName)
                combo.setCurrentIndex(ix)
            qty=query.value(3).toInt()[0]
            data=QtGui.QTableWidgetItem(self.tr("%1").arg(qty))
            self.itemTable.setItem(row,QTY,data)
            rate=query.value(4).toInt()[0]
            data=QtGui.QTableWidgetItem(self.tr("%1").arg(rate))
            self.itemTable.setItem(row,RATE,data)
            row=row+1
        self.setBuyerAddress()


class TopForm(QtGui.QDialog):
    def __init__(self,parent=None):
       super(TopForm,self).__init__(parent)
       self.setWindowTitle("Manage Invoices")
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
       self.connect(self.editWidget.viewAllInvButton,QtCore.SIGNAL('clicked()'),self.showViewWidget)
       self.connect(self.viewWidget.table,QtCore.SIGNAL('doubleClicked(QModelIndex)'),self.editButtonOfViewClicked)
       self.connect(self.viewWidget.closeButton,QtCore.SIGNAL('clicked()'),self.close)

       self.connect(self.editWidget.deleteButton,QtCore.SIGNAL('clicked()'),self.deletButtonOfEditClicked)
    def showViewWidget(self):
       self.bookWidget.setCurrentIndex(0)
       self.viewWidget.model.select()
    def deletButtonOfEditClicked(self):
       q=QtSql.QSqlQuery()
       s='DELETE  FROM `InvoicesAndProducts` WHERE `invno` = '+"\'"+str(self.savedID)+"\'"
       q.exec_(s)

       s='DELETE  FROM `Invoices` WHERE `invno` = '+"\'"+str(self.savedID)+"\'"
       q.exec_(s)
       self.editWidget.invEdit.setText("")
       self.editWidget.clearAllNow()
    def goForEdit(self):
       self.editing=True
       self.bookWidget.setCurrentIndex(1)

    def addButtonOfViewClicked(self):
       self.editing=False
       self.bookWidget.setCurrentIndex(1)

    def editButtonOfViewClicked(self):
       self.editing=True
       self.bookWidget.setCurrentIndex(1)
       for i in range(6):
           id=self.viewWidget.model.record().value(i).toString()
       index=self.viewWidget.table.currentIndex()
       #index1=QtCore.QModelIndex()
       row=index.row()
       index=self.viewWidget.model.createIndex(row,0)
       dt=self.viewWidget.model.data(index).toInt()[0]
       s='SELECT * FROM `Invoices` WHERE `invno` = '+ "\'" + str(dt) +"\'"
       q=QtSql.QSqlQuery()
       q.exec_(s)

       if (q.next()):
           id=q.value(0).toInt()[0]
           self.editWidget.invEdit.setText(str(id))
           self.savedID=id
           self.editWidget.loadInvoiceFromDB()


       edits=self.editWidget.findChildren(QtGui.QLineEdit)
       l=edits.__len__()

    def saveButtonOfEditClicked(self):
       q=QtSql.QSqlQuery()
       try:
           if (self.editing):
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

    app.exec_()