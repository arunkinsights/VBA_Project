Attribute VB_Name = "Create_Workbook_from_Data"
' This subroutine creates separate workbooks from the data on the "Branch" sheet.
' For each branch listed in column "A" of the "Branch" sheet, a new workbook is created.
' The rows corresponding to each branch (filtered by branch name) are copied,
' and pasted into a new workbook which is then saved with the branch name as the file name.

Sub Create_Workbook_from_Data()

Dim x As Long, br As String

' Turn off screen updating to speed up the process
Application.ScreenUpdating = False

' Loop through each row in the "Branch" sheet, starting from row 1 to the last non-empty row in column A.
For x = 1 To Sheets("Branch").Range("a" & Application.Rows.Count).End(xlUp).Row
    ' Assign the branch name in column A of the current row to the variable 'br'.
    br = Sheets("Branch").Range("a" & x).Value

    ' Select cell A1, then filter data based on the branch name (from column C).
    Range("A1").Select
    Selection.AutoFilter Field:=3, Criteria1:=br

    ' Select the filtered data and extend the selection to the last cell in the range (both down and to the right).
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select

    ' Copy the selected range.
    Selection.Copy

    ' Create a new workbook and paste the copied data.
    Workbooks.Add
    ActiveCell.PasteSpecial xlPasteAll

    ' Save the new workbook using the branch name as the file name in the specified directory.
    ActiveWorkbook.SaveAs "C:\Users\Arun Kumar Singh\Desktop\VBA\Created_Workbook_from_Data\" & br & ".xlsx"

    ' Close the newly created workbook.
    ActiveWorkbook.Close (True)

Next x

End Sub

