from PyQt4 import QtGui,QtCore
#from afxres import AFX_IDS_COLOR_MENUBAR
from InvoiceEntryWidget import TopForm
from SQLTableForm import SQLTableForm
from ProductForm  import ProductForm


class MainWindow(QtGui.QMainWindow):

    def __init__(self,parent):
        QtGui.QMainWindow.__init__(self,parent)

        fileMenu=self.menuBar().addMenu("&Manage")
        buyerAction=QtGui.QAction("Manage &Buyers",self)
        buyerAction.setShortcut('Ctrl+B')
        buyerAction.triggered.connect(self.manageBuyers)
        fileMenu.addAction(buyerAction)

        productAction=QtGui.QAction("Manage &Products",self)
        productAction.setShortcut('Ctrl+P')
        productAction.triggered.connect(self.manageProducts)
        fileMenu.addAction(productAction)
        f=TopForm()
        self.setCentralWidget(f)

    def manageBuyers(self):
        print "hello buyer"
        self.form=SQLTableForm()
        self.form.exec_()

    def manageProducts(self):
        print "hello product"
        self.form=ProductForm()
        self.form.exec_()
