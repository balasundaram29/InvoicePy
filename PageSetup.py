
from win32com.client import Dispatch,constants
def setupPage(sheet,typeString='Original',row=1,column=6):
    sheet.Cells(row,column).Value=typeString
    sheet.Cells(row,column).HorizontalAlignment=constants.xlRight
    page=sheet.PageSetup

    ''' page.LeftHeader = ""
    page.CenterHeader = ""
    page.RightHeader = ""
    page.LeftFooter = ""
    page.CenterFooter = ""
    page.RightFooter = ""'''
    page.LeftMargin = 0
    page.RightMargin = 0
    page.TopMargin = 0
    page.BottomMargin = 0
    page.HeaderMargin = 0
    page.FooterMargin = 0
    '''page.PrintHeadings = False
    page.PrintGridlines = False
    page.PrintComments = constants.xlPrintNoComments'''
    page.CenterHorizontally = True
    page.CenterVertically = True
    #page.Orientation = constants.xlPortrait
    page.Draft = False
    page.PaperSize = constants.xlPaperA4
    '''page.FirstPageNumber = constants.xlAutomatic
    page.Order = constants.xlDownThenOver
    page.BlackAndWhite = False
    page.Zoom = 100
    page.PrintErrors = constants.xlPrintErrorsDisplayed
    page.OddAndEvenPagesHeaderFooter = False
    page.DifferentFirstPageHeaderFooter = False
    page.ScaleWithDocHeaderFooter = True
    page.AlignMarginsHeaderFooter = True
    page.EvenPage.LeftHeader.Text = ""
    page.EvenPage.CenterHeader.Text = ""
    page.EvenPage.RightHeader.Text = ""
    page.EvenPage.LeftFooter.Text = ""
    page.EvenPage.CenterFooter.Text = ""
    page.EvenPage.RightFooter.Text = ""
    page.FirstPage.LeftHeader.Text = ""
    page.FirstPage.CenterHeader.Text = ""
    page.FirstPage.RightHeader.Text = ""
    page.FirstPage.LeftFooter.Text = ""
    page.FirstPage.CenterFooter.Text = ""
    page.FirstPage.RightFooter.Text = ""'''
