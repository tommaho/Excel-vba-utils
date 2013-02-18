Attribute VB_Name = "ModTabletoList"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'
'
'Function Name:     f_TableToList               Author:            tommaho
'Create Date:       2-18-2013                   Last Revision:
'
'Description:       Converts a multi-column excel table to a single column
'                   list.
'
'Definition:
'   f_TableToList(
'           SourceSheet as Worksheet,   <* sheet containing table
'           SourceRange as Range,       <* range to be transformed
'           DestSheet as Worksheet,     <* worksheet to place list
'           DestCell as Range           <* range to start list
'           ) as Boolean
'
'Typical Usage:
'   Private Sub cmdConvert_Click()
'       Call f_TableToList(
'           ActiveSheet
'           Selection, _
'           ActiveWorkbook.Sheets("Destination"), _
'           ActiveWorkbook.Sheets("Destination").Range("A1"))
'   End Sub
'
'Notes:
'   It's the users responsibility to make sure the destination sheet and range are
'   clear and ready to receive the transformed output
'
'
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'

Public Function f_TableToList(SourceSheet As Worksheet, _
                                SourceRange As Range, _
                                DestSheet As Worksheet, _
                                DestCell As Range) As Boolean
    On Error GoTo ErrHandler
    
    Application.ScreenUpdating = False
    
    Dim startCell, tableWidth, columnWidth, i, n
    Dim tableArray() As String
    
        Set startCell = ActiveCell
        tableWidth = SourceRange.Columns.Count
        tableHeight = SourceRange.Rows.Count
        
    ReDim tableArray(0 To tableHeight, 0 To tableWidth)
    
    'Load the array
    startCell.Activate
   
   For i = 0 To tableHeight - 1
        For n = 0 To tableWidth - 1
                tableArray(i, n) = Selection.Cells(i + 1, n + 1).Value
        Next n
   Next i
   
    startCell.Activate
    
    'Unload the array
    DestSheet.Activate
    DestCell.Activate
        
    For i = 1 To tableHeight - 1
        For n = 1 To tableWidth - 1
            ActiveCell.Value = tableArray(i, 0)
            ActiveCell.Offset(0, 1).Value = tableArray(0, n)
            ActiveCell.Offset(0, 2).Value = tableArray(i, n)
            ActiveCell.Offset(1, 0).Activate
        Next n
    Next i
    
    Erase tableArray
    
    DestCell.Activate
    SourceSheet.Activate
    startCell.Activate
    
    f_TableToList = True
    Application.ScreenUpdating = True
    
    Exit Function
    
ErrHandler:
    MsgBox ("An error has occurred: " & Err.Description)
    Erase tableArray
    f_TableToList = False
    Application.ScreenUpdating = True
End Function
