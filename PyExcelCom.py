#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      ANNAIENG
#
# Created:     19/02/2014
# Copyright:   (c) ANNAIENG 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from win32com.client import Dispatch,constants
XL_EDGE_TOP=10
XL_CONT_LINE=1

import time

    START_ROW=1
    START_COL=1
    #we will be thinking in terms of current row and column
    CURR_ROW=1
    CURR_COL=1
    FINAL_COL=9
    INVNO_COL=7
    INV_DATE_COL=7
    SNO_COL=CURR_COL
    DESCPN_COL=SNO_COL+1
    QUANTITY_COL=DESCPN_COL+5
    RATE_COL=QUANTITY_COL+1
    AMOUNT_COL=RATE_COL+1
    excelApp=Dispatch('Excel.Application')
    excelApp.Visible=True
    book=excelApp.Workbooks.Add()
    sheet=book.Worksheets(1)

    range=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
    range.ColumnWidth=8
    range=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
    range.ColumnWidth=12
    range=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
    range.ColumnWidth=16

    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL))
    range.Merge()
    range.Font.Size=22
    sheet.Cells(CURR_ROW,CURR_COL).Value="INVOICE"
    range.HorizontalAlignment=constants.xlCenter
    range.VerticalAlignment=constants.xlTop
    range.Borders(constants.xlEdgeTop).LineStyle=constants.xlContinuous

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=16
    sheet.Cells(CURR_ROW,CURR_COL).Value="Christech Engineering Company"


    range=sheet.Range(sheet.Cells(CURR_ROW,INVNO_COL),sheet.Cells(CURR_ROW,INVNO_COL+1))
    range.Merge()
    range.Font.Size=16
    sheet.Cells(CURR_ROW,INVNO_COL).Value="Invoice No :"


    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=16
    sheet.Cells(CURR_ROW,CURR_COL).Value="JV Nagar"

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=16
    sheet.Cells(CURR_ROW,CURR_COL).Value="VK Road,Vinayagapuram"

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=16
    sheet.Cells(CURR_ROW,CURR_COL).Value="Coimbatore12"

    CURR_ROW=CURR_ROW+1
    TINNO_ROW=CURR_ROW
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
    range.Merge()
    range.Font.Size=14
    sheet.Cells(TINNO_ROW,CURR_COL).Value="TIN No :"
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=16
    CURR_ROW=CURR_ROW+1
    sheet.Cells(TINNO_ROW,CURR_COL).Value="Buyer's Name and Address :"
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=12

    sheet.Cells(CURR_ROW,CURR_COL).Value="bUYEREering Company"

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=13
    sheet.Cells(CURR_ROW,CURR_COL).Value="bUEYR ADD1"

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=12
    sheet.Cells(CURR_ROW,CURR_COL).Value="Buyer add2"

    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,FINAL_COL-3))
    range.Merge()
    range.Font.Size=12
    sheet.Cells(CURR_ROW,CURR_COL).Value="Buyer add3"

    CURR_ROW=CURR_ROW+1
    TINNO_ROW=CURR_ROW
    range=sheet.Range(sheet.Cells(CURR_ROW,CURR_COL),sheet.Cells(CURR_ROW,CURR_COL+3))
    range.Merge()
    range.Font.Size=12
    sheet.Cells(TINNO_ROW,CURR_COL).Value="Buyers TIN No :"



    CURR_ROW=CURR_ROW+2
    BUYER_ROW=CURR_ROW

    range=sheet.Range(sheet.Cells(CURR_ROW,SNO_COL),sheet.Cells(CURR_ROW,SNO_COL))
    #range.Merge()
    range.Font.Size=12
    range.ColumnWidth=4
    sheet.Cells(CURR_ROW,SNO_COL).Value="SNo"

    range=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW,QUANTITY_COL-1))
    range.Merge()
    range.Font.Size=12
    range.HorizontalAlignment=constants.xlCenter
    sheet.Cells(CURR_ROW,DESCPN_COL).Value="Description"
    range=sheet.Range(sheet.Cells(CURR_ROW,QUANTITY_COL),sheet.Cells(CURR_ROW,QUANTITY_COL))
    range.Merge()
    range.Font.Size=12
    range.HorizontalAlignment=constants.xlCenter
    sheet.Cells(CURR_ROW,QUANTITY_COL).Value="Quantity"

    range=sheet.Range(sheet.Cells(CURR_ROW,RATE_COL),sheet.Cells(CURR_ROW,RATE_COL))
    range.Merge()
    range.Font.Size=12
    range.HorizontalAlignment=constants.xlCenter
    sheet.Cells(CURR_ROW,RATE_COL).Value="Rate"

    range=sheet.Range(sheet.Cells(CURR_ROW,AMOUNT_COL),sheet.Cells(CURR_ROW,AMOUNT_COL))
    range.Merge()
    range.Font.Size=12
    range.HorizontalAlignment=constants.xlCenter
    sheet.Cells(CURR_ROW,AMOUNT_COL).Value="Amount"

    ROWS_PER_ITEM=3
    CURR_ROW=CURR_ROW+1
    range=sheet.Range(sheet.Cells(CURR_ROW,DESCPN_COL),sheet.Cells(CURR_ROW+ROWS_PER_ITEM-1,QUANTITY_COL-1))
    range.Merge()
    range.VerticalAlignment = constants.xlTop
    range.WrapText = True



    #ime.sleep(20)
if __name__ == '__main__':
    main()
