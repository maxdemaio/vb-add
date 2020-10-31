Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
' Maxwell G. DeMaio          '''
' vb-add                     '''
' Data Clean Module          '''
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''


Public Function isValid(strInput As String) As Boolean
''' Public function to validate if string is a valid column
''' Default value is False

    On Error GoTo isValid_Error

    Dim rngSet  As Range
    Set rngSet = Range(strInput & "1")
    isValid = True

    On Error GoTo 0
    Exit Function

isValid_Error:
End Function



Sub emptyRowClean()
''' Get rid of the rows that are completely empty

Dim currSheetNum As Double
Dim LastRow As Long
Dim i As Long

LastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row

' Make sure original data is safe
On Error GoTo ErrorHandler
ActiveSheet.Copy After:=ActiveWorkbook.ActiveSheet

For i = LastRow To 1 Step -1
    If Application.CountA(Rows(i)) = 0 Then
        Rows(i).Delete
    End If
Next

Exit Sub

' If namespace error occurs, add data to Sheet(x) and perform subroutine
' Switch active sheet to Sheet(x)
ErrorHandler:
    currSheetNum = ActiveSheet.Index
    Sheets(currSheetNum).Cells.Copy Destination:=Sheets(currSheetNum + 1).Range("A1")
    Sheets(currSheetNum + 1).Activate
    Resume Next

End Sub

Sub columnClean()
''' Get rid of the columns that are completely empty

Dim currSheetNum As Double
Dim i As Integer

' Make sure original data is safe
On Error GoTo ErrorHandler
ActiveSheet.Copy After:=ActiveWorkbook.ActiveSheet

i = ActiveSheet.Cells.SpecialCells(xlLastCell).Column

Do Until i = 0
    If WorksheetFunction.CountA(Columns(i)) = 0 Then
        Columns(i).Delete
    End If
i = i - 1
Loop

Exit Sub

' If namespace error occurs, add data to Sheet(x) and perform subroutine
' Switch active sheet to Sheet(x)
ErrorHandler:
    currSheetNum = ActiveSheet.Index
    Sheets(currSheetNum).Cells.Copy Destination:=Sheets(currSheetNum + 1).Range("A1")
    Sheets(currSheetNum + 1).Activate
    Resume Next
    
End Sub

Sub unhideRowsAndColumns()
''' Unhide all rows and columns

Rows.Hidden = False

Columns.Hidden = False

End Sub

Sub autoFit()
''' Autofit every cell in the active worksheet

ActiveSheet.Columns.autoFit
ActiveSheet.Rows.autoFit

End Sub

Sub rowSpecifyClean()
''' Delete rows based on if a specific cell is empty in a specific column
''' If the cell has data, skips to next row

Dim myColumn As String
Dim currSheetNum As Double

' Get column from user input
On Error Resume Next
    myColumn = Application.InputBox( _
      Title:="Row Specify Clean", _
      Prompt:="Select a column to scan for empty cells, deleting the rows where they occur (Ex. A, B, C, etc.)", _
      Type:=2)
    On Error GoTo 0

' Test to ensure User Did not cancel
If myColumn = "" Or myColumn = "False" Then Exit Sub

' Make sure original data is safe
On Error GoTo ErrorHandler
ActiveSheet.Copy After:=ActiveWorkbook.ActiveSheet

On Error Resume Next
Columns(myColumn).SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Exit Sub

' If namespace error occurs, add data to Sheet(x) and perform subroutine
' Switch active sheet to Sheet(x)
ErrorHandler:
    currSheetNum = ActiveSheet.Index
    Sheets(currSheetNum).Cells.Copy Destination:=Sheets(currSheetNum + 1).Range("A1")
    Sheets(currSheetNum + 1).Activate
    Resume Next
    
End Sub

Sub nameDragDown()
''' Name Drag Down
''' Obtain range from user and drag down strings

Dim company As String
Dim i As Double
Dim myStart As Double
Dim myEnd As Double
Dim myColumn As String
Dim currSheetNum As Double

' Three forms column, start row, end row,
' Get column from user input
On Error Resume Next
    myColumn = Application.InputBox( _
      Title:="Name Drag Down", _
      Prompt:="Select a column where names should be dragged down (Ex. A, B, C, etc.)", _
      Type:=2)
    On Error GoTo 0
    
' Test to ensure User Did not cancel
If myColumn = "" Or myColumn = "False" Then Exit Sub

' Make sure column string is valid
If isValid(myColumn) = False Then Exit Sub

' Get starting row from user input
On Error Resume Next
    myStart = Application.InputBox( _
      Title:="Name Drag Down", _
      Prompt:="Select a row to start dragging down (Ex. 1, 2, 3, etc.)", _
      Type:=1)
    On Error GoTo 0
    
' Test to ensure User Did not cancel
If myStart = False Then Exit Sub

' Get ending row from user input
On Error Resume Next
    myEnd = Application.InputBox( _
      Title:="Name Drag Down", _
      Prompt:="Select a row to stop dragging down (Ex. 1, 2, 3, etc.)", _
      Type:=1)
    On Error GoTo 0
    
' Test to ensure User Did not cancel
If myEnd = False Then Exit Sub

' Check to make sure there are no negative values
If myStart < 0 Or myEnd < 0 Then
    Exit Sub
End If

' Make sure original data is safe
On Error GoTo ErrorHandler
ActiveSheet.Copy After:=ActiveWorkbook.ActiveSheet
    
For i = myStart To myEnd ' All of the data
    If IsEmpty(Cells(i, myColumn)) = False Then ' If there is an account name
        company = Cells(i, myColumn).Value  ' Set company equal to it
    Else
        Cells(i, myColumn).Value = company ' Else, there's no account name, set it to the account name
    End If
Next i

Exit Sub

' If namespace error occurs, add data to Sheet(x) and perform subroutine
' Switch active sheet to Sheet(x)
ErrorHandler:
    currSheetNum = ActiveSheet.Index
    Sheets(currSheetNum).Cells.Copy Destination:=Sheets(currSheetNum + 1).Range("A1")
    Sheets(currSheetNum + 1).Activate
    Resume Next
    
End Sub

Sub rowSpecifyContentsClean()
''' Get rid of the rows that contain a specific string

Dim LastRow As Long
Dim i As Long
Dim myColumn As String
Dim myString As String
Dim currSheetNum As Double

'' Two forms to obtain input
' Get string from user input
On Error Resume Next
    myString = Application.InputBox( _
      Title:="Row Specify Contents Clean", _
      Prompt:="Specify a string to search for, deleting the rows where it occurs", _
      Type:=2)
    On Error GoTo 0

' Test to ensure User Did not cancel
If myString = "" Or myString = "False" Then Exit Sub


' Get column from user input
On Error Resume Next
    myColumn = Application.InputBox( _
      Title:="Row Specify Contents Clean", _
      Prompt:="Select a column to scan for cells with the specified string (Ex. A, B, C, etc.), deleting the rows where it occurs", _
      Type:=2)
    On Error GoTo 0
    
' Test to ensure User Did not cancel
If myColumn = "" Or myColumn = "False" Then Exit Sub

' Make sure column string is valid
If isValid(myColumn) = False Then Exit Sub


'' Delete rows with the specified string
' Make sure original data is safe
On Error GoTo ErrorHandler
ActiveSheet.Copy After:=ActiveWorkbook.ActiveSheet

LastRow = ActiveSheet.Cells.SpecialCells(xlLastCell).Row
For i = LastRow To 2 Step -1
    If Application.Cells(i, myColumn).Value = myString Then
        Rows(i).Delete
    End If
Next

Exit Sub

' If namespace error occurs, add data to Sheet(x) and perform subroutine
' Switch active sheet to Sheet(x)
ErrorHandler:
    currSheetNum = ActiveSheet.Index
    Sheets(currSheetNum).Cells.Copy Destination:=Sheets(currSheetNum + 1).Range("A1")
    Sheets(currSheetNum + 1).Activate
    Resume Next

End Sub

