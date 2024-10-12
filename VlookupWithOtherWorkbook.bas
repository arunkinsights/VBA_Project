Attribute VB_Name = "VlookupWithOtherWorkbook"
Option Explicit
Option Base 1

Sub VlookupWithDynamicArrayFromDifferentWorkbook()
  Dim darray() As Variant
  Dim x As Long, y As Long, lr As Long, lc As Long
  Dim sourceWorkbook As Workbook
  Dim sourceSheet As Worksheet
  Dim destinationSheet As Worksheet

  ' Open the source workbook (ensure the path is correct)
  Set sourceWorkbook = Workbooks.Open("C:\Users\Arun Kumar Singh\Desktop\Dump\RawData.xlsx")
  Set sourceSheet = sourceWorkbook.Sheets("data")

  ' Find the last row and last column with data in the source workbook
  lr = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row  ' Use Cells(Rows.Count, 1) instead of Range("A")
  lc = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column

  ' Resize the array based on the last row and column in the source workbook
  ReDim darray(1 To lr, 1 To lc)

  ' Loop through the rows and columns and populate the array from the source workbook
  For x = 1 To lr
    For y = 1 To lc
      darray(x, y) = sourceSheet.Cells(x, y).Value
    Next y
  Next x

  ' Define the destination sheet (replace with your desired location)
  Set destinationSheet = ThisWorkbook.Sheets("Sheet1")  ' Change "Sheet1" to your actual sheet name

  ' Loop through each row in the destination sheet and perform VLOOKUP
  Dim v As Long, lookupValue As String
  For v = 2 To destinationSheet.UsedRange.Rows.Count
    lookupValue = destinationSheet.Cells(v, 1).Value  ' Assuming lookup value is in column A
    destinationSheet.Cells(v, 2).Value = Application.WorksheetFunction.VLookup(lookupValue, darray, 3, False) ' Use False for exact match
  Next v

  ' Close the source workbook (Close SaveChanges argument is optional)
  sourceWorkbook.Close SaveChanges:=False

End Sub
