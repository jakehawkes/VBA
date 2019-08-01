Attribute VB_Name = "CopyTemplate"
' Specially designed to copy the table from the "Template" tab,  into the "Output" tab.
' to the output of the Equipment Schedule for the AT RWIS Expansion project
'
' Copies from A1 in the Template tab to A1 in the Output tab
' Copies from A1 down to wherever "END OF TEMPLATE" is located, in whichever column it is in
' Copies all formatting values and column widths. 
'
' Change History
' Author    Date        Change
' --------  ----------  --------------------------------------------
' Jake      June        Written
' Jake      2012-07-09  Added output tab protection code
'
'
Sub CopyTemplate()
    Dim srcRange As Range
    Dim srcCell As Range
    Dim srcEndMarker As Range
    Dim destCell As Range
    Dim ws_src As Worksheet
    Dim ws_dest As Worksheet
    Dim i As Integer
    
        
    Set ws_src = Worksheets("Template")
    Set ws_dest = Worksheets("Output")
    
    Set srcEndMarker = ws_src.Cells.Find("END OF TEMPLATE")
        
    Application.ScreenUpdating = False
    ws_dest.Unprotect

    With ws_src
        Set srcRange = .Range(.Cells(1, 1), .Cells(srcEndMarker.Row - 1, srcEndMarker.Column))
    End With
    With ws_dest
        'delete the output area
        .Range("A1:G100").Clear
        ' now set the destination range
        Set destCell = .Range("A1")
    End With
    
    For Each srcCell In srcRange.Rows
        ' Carefull: the default value of an excel cell is 'Empty' so you must
        '           check with IsEmpty() since .value = 0 and .value = empty are equivalent
        If srcCell.Cells(1, 6).Value <> 0 Or IsEmpty(srcCell.Cells(1, 6)) Then
            ' cell is empty or non-zero, copy the row
            srcCell.Copy
            destCell.PasteSpecial (xlPasteAll)
            destCell.PasteSpecial (xlPasteColumnWidths)
            Set destCell = destCell.Offset(1, 0)
        End If
        'if you want to dump out of the 'for each' loop, use 'exit for'
    Next srcCell

    Application.ScreenUpdating = True
    ws_dest.Protect
    
End Sub

