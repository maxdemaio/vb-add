Attribute VB_Name = "Module2"
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''
' Maxwell G. DeMaio          '''
' vb-add                     '''
' Information Module         '''
''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''

Sub displayInfo()
''' vb-add all subroutines information

MsgBox "Delete Empty Columns" & vbNewLine & "Deletes any empty columns present in the active worksheet." & vbNewLine & _
    vbNewLine & _
    "Delete Empty Rows" & vbNewLine & "Deletes any empty rows present in the active worksheet." & vbNewLine & _
    vbNewLine & _
    "Row Specify Deletion" & vbNewLine & "After specifying a column, deletes any rows where there is an empty cell." & vbNewLine & _
    vbNewLine & _
    "Auto Fit All" & vbNewLine & "Auto fits all cells in the active worksheet." & vbNewLine & _
    vbNewLine & _
    "Unhide All Rows and Columns" & vbNewLine & "Unhides all cells in the active worksheet." & vbNewLine & _
    vbNewLine & _
    "Row String Specify Deletion" & vbNewLine & "After specifying a string and a column, deletes any rows where the string occurs." & vbNewLine & _
    vbNewLine & _
    "Name Drag Down" & vbNewLine & "Given a column, starting and ending row, the subroutine will drag names down." & vbNewLine _
    , vbInformation, "vb-add Information"
    
End Sub

