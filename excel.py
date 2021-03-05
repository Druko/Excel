
#!/usr/bin/env python3.7

import win32com.client.dynamic

import os
import string

try:
    import wx #wxpython
except ImportError:
    raise ImportError("The wx module (wxpython) is required.")

try:
    import numpy as np
except ImportError:
    raise ImportError("The numpy module is required.")



class Excel(object):
    '''
    Excel Interface with Numpy helper module
    '''

    def __init__(self, visible = False):
        """
        Initialize the COM interface
        after this _init_ call Open(filename)
        note: 
        Sometimes this will fail therefore go to task manager and end task excelapp.
        """
        self._excel = win32com.client.dynamic.Dispatch("Excel.application")
        self._excel.Visible = visible
        self._excel_file_path = None
        self._sheetName = None
        self._sheet_values = np.array(0)        


    def GetNumberOfSheet(self):
        '''
        Return the number of Sheets

        its requiere a Workbooks.Open
        '''
        return self._excel.ActiveWorkbook.Sheets.Count
        
    def GetSheetNames(self):
        """
        Return a list of sheet in the current workbook.

        note:
        its requiere a Workbooks.Open
        """
        numOfSheets = self.GetNumberOfSheet()
        return [self._excel.Sheets(index).Name for index in range(1,numOfSheets+1)]


    def __del__(self):
        """
        
        """        
        if not self._excel.Visible:
            print
            print ('-'*50)
            print ("\nClose and Quit excel")
            print ('-'*50)
            self.close()
            self.quit()
            del self._excel
        else:
            del self._excel


    def close(self):
        '''
        close workbook
        '''
        if self._excel:
            self._excel.Workbooks.Close()

    def quit(self):
        '''
        close excel app
        '''
        if self._excel:
            self._excel.Application.Quit()


    
    @staticmethod
    def fileopenbox(msg = ''):
        """
        return a string else exit
        """
        #wildcard = "Excel files (*.xls;*.xlsx) |*.xls;*.xlsx;|All Files (*.*)|*.*"
        wildcard = "Excel files (*.xls;*.xlsx) |*.xls;*.xlsx;"
        gui = wx.App()        
        dialog = wx.FileDialog(None, msg ,'', '', wildcard, wx.FD_OPEN | wx.FD_CHANGE_DIR)
        dialog.ShowModal()
        path = dialog.GetPath()
        dialog.Destroy()        
        return path if path != '' else exit()

    @staticmethod
    def ChoiceBox(lst, title='' ,msg = ''):
        """
        return a string
        """
        gui   = wx.App()
        style = wx.OK | wx.CANCEL | wx.CENTRE | wx.DEFAULT_DIALOG_STYLE
        dialog = wx.SingleChoiceDialog(None, msg, title, list(lst), style )
        dialog.ShowModal()
        selection = dialog.GetStringSelection()
        dialog.Destroy()
        return selection if selection != '' else exit()

    @staticmethod
    def MultiChoiceBox(lst, title='' ,msg = ''):
        """
        returns a list of index for selections
        """
        gui   = wx.App()
        style = wx.OK | wx.CANCEL | wx.CENTRE | wx.DEFAULT_DIALOG_STYLE
        dialog = wx.MultiChoiceDialog(None, msg, title, list(lst), style )
        dialog.ShowModal()
        selection = dialog.GetSelections()   
        dialog.Destroy()
        return selection if selection != '' else exit()


    def GetFilename(self):
        return os.path.basename(self._excel_file_path)

    def Open(self, filename):
        """
        Open an existing Excel workbook.
        return a workbook  
        """
        self._excel_file_path = filename
        print('Filename: %s' % self.GetFilename())
        return self._excel.Workbooks.Open(filename)

    def _ReadSheet(self, sheetName = None):
        """
        Read all values from a single sheet
        container numpy.array (_sheet_values)
       
        return: None

        Note: 
        Excel class was intended for use for inheritence
        python does not have virtual keyword therefore I decided to use the underscore (_ReadSheet())
        example:
        def _ReadSheet(sheetName): #derived override
            super(derived, self)._ReadSheet(sheetName) #calling base clase
            #TODO: derive implementaion...
        """
        try:
            wb_sheet = self._excel.Sheets(sheetName)
            wb_sheet.Activate()
            print ("SheetsName %s" % (wb_sheet.Name))
        except:
            raise LookupError('SheetsName %s not found' % (sheetName))

        last_cell = self._excel.ActiveSheet.Cells.SpecialCells(11) #xLastCell = 11  "Get the location of the last_cell (Equivalent to Ctrl+End)"
        
        #example: (np.array Reference subscript access) 
        #arr[1][1:] #Get all info in row 2 starting at col 'B'
        #arr[:,1]   #Row start at index 0 and Get all info from col index 1
        #arr[1:,0]  #Row start at index 1 and Get all info from col index 0
        self._sheet_values = np.array(wb_sheet.Range(wb_sheet.Cells(1, 1), last_cell).Value2 , dtype= str) #Get all the value from (Row 1, col A) to the last_cell dtype= string

    def GetSheetValues(self):
        '''
        return: numpy.array
        '''
        return self._sheet_values

    def ReadSheetGetValues(self,sheetName):
        self._ReadSheet(sheetName)
        return self.GetSheetValues()


    def find_cell_index_loc(self,cellValue):
        """
        Search and return 2d index location
        in: cellValue as a string
        return a tuple of array (row,col) indexs location
        """
        if isinstance(cellValue, str ):            
            try:
                index = np.nonzero(self._sheet_values == cellValue)

                if index[0].size and index[1].size:
                    row,col = tuple(index)
                    print ('index [row:%s][col:%s] = %s' %(row,col,self._sheet_values[index]))
                    return (row,col)
                else:
                    raise LookupError ('string "%s" not found' %(cellValue)) 

            except Exception as e:
                print ( "error: ",e.args[0]) 
        else:
            raise TypeError('argumet must be a string')                
              
    
    @staticmethod
    def col_num2char(n):
        '''
        Helper Function
        Return corresponding  col letter base on Number
        example:
        GetColumnName(1)  #return A
        GetColumnName(3)  #return C
        GetColumnName(27)  #return AA
        '''
        MAX = 50
 
        # To store result (Excel column name)
        char = ["\0"]*MAX
 
        # To store current index in str which is result
        i = 0
 
        while n > 0:
            # Find remainder
            rem = n%26
 
            # if remainder is 0, then a 'Z' must be there in output
            if rem == 0:
                char[i] = 'Z'
                i += 1
                n = (n/26)-1
            else:
                char[i] = chr((rem-1) + ord('A'))
                i += 1
                n = n/26
        char[i] = '\0'
 
        # Reverse the string and print result
        char = char[::-1]
        return  "".join(char).strip('\0')

    @staticmethod
    def col_char2num(col):
        """
        Helper Function
        Return the number representation of the corresponding  col letter
        example:
        col2num('AA') <-- return 27        
        """   
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        return num

    def GetTotalCol(self):
        return len(self._sheet_values[0]) #_sheet_values[row][col]

    def GetTotalRow(self):
        return len(self._sheet_values)#_sheet_values[row][col] 

    
    def sort_col(self,col_start,row_num):
        '''
        Return a list of index location for sorting the array
        '''
        if  isinstance(col_start,(unicode,str )):
            col_start = self.col_char2num(col_start) - 1
            row_num -= 1

        #example:
        #print arr[1][1:] #Get all info in row 2 starting at col'B'
        return np.argsort(self._sheet_values[row_num][col_start:])

    def sort_row(self,col , row_start):
        '''
        Return a list of index location for sorting the array
        '''
        if  isinstance(col,str):
            col = self.col_char2num(col_num) - 1
            row_start -= 1
        #example:
        #print arr[2:,1] #Row start at row 3 and get all info form col 'B'
        #print arr[:,0]  #get all info col 'A' 
        #print arr[1][1:][0] #start at row index 1 starting at col index1 , the 3rd subscript is to access the data       
        return np.argsort(self._sheet_values[row_start:,col])

    def GetallColfromRow(self,col_start = 'A', row_num = 1):
        
        if  isinstance(col_start,str):
            col_start = self.col_char2num(col_start) - 1
            row_num -= 1           

        return self._sheet_values[row_num][col_start:]


    def GetallRowfromCol(self,col = 'A',row_start = 1):

        if  isinstance(col,str):
            col = self.col_char2num(col) - 1
            row_start -=1

        return self ._sheet_values[row_start:,col]


  
    


def main():
    '''
    DEMO
    '''
    print ("Excel interface demo")

    path =Excel.fileopenbox("select any excel file")
    
    xls = Excel(True)
    xls.Open(path)
    
    count = xls.GetNumberOfSheet()
    print("number of sheet", count)

    sheets = xls.GetSheetNames();
    print("sheets names", sheets)

    xls._ReadSheet(sheets[0]) 
       
    location = xls.find_cell_index_loc('decimal')

    values = xls.GetSheetValues()

    print (values[location])

    col = xls.GetallColfromRow();
    print(col)

    row = xls.GetallRowfromCol()
    print(row)

    print()

    


if __name__ == '__main__':
    main()
     