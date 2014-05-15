from PyQt4 import QtGui
class InvoiceTableWidget(QtGui.QTableWidget):
    def __init__(self):
        super(InvoiceTableWidget,self).__init__()

    def sizeHint(self):
        return (self.rowCount()*50)+50
