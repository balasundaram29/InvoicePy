#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      ANNAIENG
#
# Created:     11/02/2014
# Copyright:   (c) ANNAIENG 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from PyQt4 import QtGui,QtCore
import string
numEdit=0
wordEdit=0
alphaDict={0:'zero',1:'one',2:'two',3:'three',4:'four',5:'five',6:'six',7:'seven',8:'eight',9:'nine',
               10:'ten',11:'eleven',12:'twelve',13:'thirteen',14:'fourteen',15:'fifteen',16:'sixteen',
               17:'seventeen',18:'eighteen',19:'nineteen',20:'twenty',30:'thirty',40:'forty',50:'fifty',
               60:'sixty',70:'seventy',80:'eighty',90:'ninety' }
def main():
    global numEdit,wordEdit
    app=QtGui.QApplication([])
    wid=QtGui.QWidget()
    layout=QtGui.QVBoxLayout()
    numEdit=QtGui.QLineEdit()
    wordEdit=QtGui.QLineEdit()
    btn=QtGui.QPushButton("Get")
    layout.addWidget(numEdit)
    layout.addWidget(btn)
    layout.addWidget(wordEdit)
    wid.setLayout(layout)
    QtCore.QObject.connect(btn,QtCore.SIGNAL('clicked()'),convertToWords)
    wid.show()
    return app.exec_()
def convertToWords():
    n=long(numEdit.text())
    s=changeAnywhere(n)
    wordEdit.setText(s)
def changeWithinHundred(n):
    global numEdit,wordEdit,alphaDict


    if(n<=20):
        s=alphaDict[n]

    else:

        tens,ones=divmod(n,10)



        big=tens*10
        s=alphaDict[big]+" "+alphaDict[ones]

    s=" and "+s	
    return s


def changeWithinThousand(n):
   # print "n is ",n

    #n=numEdit.text().toInt()[0]
    hundreds,remainder=divmod(n,100)
    s=''
    if(hundreds>0):
        s=changeWithinHundred(hundreds)+' '+'hundred '
        if(remainder!=0):
            s=s+changeWithinHundred(remainder)

    else:

         s=changeWithinHundred(n)
    #wordEdit.setText(s)
    return s
def changeWithinLakh(n):

    thousands,remainder=divmod(n,1000)
   # print 'thousands = ',thousands

    s=' '
    if(thousands>0):
        s=changeWithinHundred(thousands)+' '+"  "+"thousand "
        if(remainder!=0):
            s=s+changeWithinThousand(remainder)
    else:
        s=changeWithinThousand(n)
    return s

def changeWithinCrore(n):
    lakhs,remainder=divmod(n,100000)
    s=' '
    if(lakhs>0):
        s=changeWithinHundred(lakhs)+' '+'lakh '
        if(remainder!=0):
            s=s+changeWithinLakh(remainder)
    else:
        s=changeWithinLakh(n)
    return s
def changeAnywhere(n):
    #n=numEdit.text().toInt()[0]
    s=''
    crores,remainder=divmod(n,10000000L)

    if(crores>0):
        s=changeAnywhere(crores)+' '+'crore '
        if(remainder!=0):
            s=s+changeWithinCrore(remainder)
    else:
        s=changeWithinCrore(n)
    
    s=string.replace(s,' and',' ',s.count(' and')-1)
    s=string.strip(s)
    s=string.replace(s,"  "," ")
    sList=string.split(s," ")
    cList=[]
    for word in sList:
       word=string.capitalize(word)

       cList.append(word)

    s = string.join(cList," ")
    s=string.strip(s)
    if( string.find(s,'And')==0):
      s=string.replace(s,"And","",1)
      s=string.strip(s)		 			
    return s
#ordEdit.setText(s)

if __name__ == '__main__':
    main()