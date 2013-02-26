Attribute VB_Name = "mod_TabletoList"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'
'
'Function Name:     f_TableToList               Author:            tommaho
'Create Date:       2-18-2013                   Last Revision:      2-25-2013
'
'Description:       Converts a multi-column excel table to a single column
'                   list.
'
'Definition:
'   f_TableToList(
'           Optional SourceSheet as Worksheet,   <* sheet containing table,
'                                                   defaults to activesheet
'           Optional SourceRange as Range,       <* range to be transformed
'                                                   defaults to selection
'           Optional DestSheet as Worksheet,     <* worksheet to place list,
'                                                   defaults to a new sheet
'           Optional DestCell as Range           <* range to start list,
'                                                   defaults to A1
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
'   2-25-13 Revision: made all arguments optional with defaults
'
'
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'

Public Function f_TabletoList(Optional SourceSheet As Variant, _
                                Optional SourceRange As Variant, _
                                Optional DestSheet As Variant, _
                                Optional DestCell As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsMissing(SourceSheet) = True Then
        Set SourceSheet = ActiveSheet
    End If
    
    If IsMissing(SourceRange) = True Then
        Set SourceRange = Selection
    End If
    
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
    
    If IsMissing(DestSheet) = True Then
        Set DestSheet = Sheets.Add
    End If
    
    If IsMissing(DestCell) = True Then
        Set DestCell = DestSheet.Range("A1")
    End If
    
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
    
    f_TabletoList = True
    Application.ScreenUpdating = True
    
    Exit Function
    
ErrHandler:
    MsgBox ("An error has occurred: " & Err.Description)
    Erase tableArray
    f_TabletoList = False
    Application.ScreenUpdating = True
End Function


