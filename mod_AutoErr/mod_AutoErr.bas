Attribute VB_Name = "mod_AutoErr"
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'
'
'Function Name:     f_AutoErr                   Author:            tommaho
'Create Date:       2-25-2013                   Last Revision:
'
'Description:       Replaces a cell's formula with an if(iserror(X),Y,X)
'                   version of itself. Works across a range or selection.
'
'Definition:
'   f_AutoErr(
'           Optional theRange As Variant,       <* range that requires
'                                               error formulas. Defaults
'                                               to selection.
'           Optional displayErrAs As Variant    <* what the formula should
'                                               display if an error is
'                                               encountered. Defaults to 0.
'           ) as Boolean
'
'Typical Usage:
'   Private Sub cmdAddErrorCheck_Click()
'       Call f_AutoErr(Selection, "Not Found")
'   End Sub
'
'Notes:
'
'VBA type identifiers for modification of displayErrAs case statement:
'
'Value     Variant type
'0     Empty (unitialized)
'1     Null (no valid data)
'2     Integer
'3     Long Integer
'4     Single
'5     Double
'6     Currency
'7     Date
'8     String
'9     Object
'10     Error Value
'11     Boolean
'12     Variant (only used with arrays of variants)
'13     Data access object
'14     Decimal value
'17     Byte
'36     User Defined Type
'8192           Array
''
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>'

Public Function f_AutoErr(Optional theRange As Variant, Optional displayErrAs As Variant) As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
        
    If IsMissing(theRange) = True Then
        theRange = Selection
    End If
    
    If IsMissing(displayErrAs) = True Then
        displayErrAs = 0
    Else
        Select Case VarType(displayErrAs)
            Case 8: 'String
                displayErrAs = Chr(34) & displayErrAs & Chr(34)
        End Select
        
    End If
    
    If Not theRange.Find(What:="ERROR", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False) Is Nothing Then
        Err.Raise 513, "ErrtoZero", "Error checking already exists in this selection."
    End If
    
    Dim sourceFormula, sourceAddress, i, selectionCount, destFormula, subFormula

    selectionCount = Selection.Count
    
    For i = 1 To selectionCount
        sourceAddress = Selection(i).AddressLocal
        If Range(sourceAddress).HasFormula = True Then
            sourceFormula = CStr(Range(sourceAddress).Formula)
            subFormula = Right(sourceFormula, Len(sourceFormula) - 1)
            destFormula = "=if(iserror(" & subFormula & ")," & displayErrAs & "," & subFormula & ")"
            Range(sourceAddress).Formula = destFormula
        End If
    Next i
    
    
    f_ErrToZero = True
    Application.ScreenUpdating = True
Exit Function
ErrHandler:
        MsgBox ("An error occurred: " & Err.Description)
        f_ErrToZero = False
        Application.ScreenUpdating = True
End Function





