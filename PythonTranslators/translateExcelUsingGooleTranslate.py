# coding: euc_jp
import win32com.client as win32
import os
import codecs


def iterateGroup (shape):
    for subshape in shape.GroupItems:
        if subshape.TextFrame2.HasText:
            if subshape.Type == 6:    # Grouped item
                iterateGroup(subshape)
            else:
                txtJP = (subshape.TextFrame2.TextRange.Text)
                print (subshape.Name + ': ')
                print (txtJP)

def openWorkbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            xlwb = xlapp.Workbooks.Open(xlfile)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)

try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    path = os.getcwd()
    print(path + '\\' + 'Test.xlsx')
    wb = openWorkbook(excel, path + '/' + 'inORG.xlsx')
    #ws = wb.Worksheets('abc') 
    excel.Visible = True

    ws = wb.Worksheets
   

    for w in ws:
        for row in w.UsedRange.Value:
            for item in row:
                if item != None:
                    print (item)
           
    print ('\nSHAPES\n')
    for w in ws:
        print w.Name
        w.Activate
        canvas = w.Shapes
        for shape in canvas:
            if shape.TextFrame.Characters:                
                if shape.Type == 4:    # Comment
                    pass
                elif shape.Type == 6:    # Grouped item
                    iterateGroup(shape)
                elif shape.TextFrame2.HasText:
                    txtJP = (shape.TextFrame2.TextRange.Text)
                    print (shape.Name + ': ')
                    print (txtJP)
                    
except Exception as e:
    print(e)

finally:
    # RELEASES RESOURCES
    ws = None
    wb = None
    excel = None