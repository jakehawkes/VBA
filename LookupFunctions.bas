Attribute VB_Name = "LookupFunctions"

Function GetDocProps(propName As String) As String
    On Error GoTo errValue
    
    ' propName can be the name of any DocumentProperty object
    ' such as: "Creation Date", "Last Save Time", "Last Print Date" etc
    GetDocProps = ActiveWorkbook.BuiltinDocumentProperties(propName)
    Exit Function
    
errValue:
    GetDocProps = CVErr(xlErrValue)
End Function

' joins with spaces the contents of all the cells in the range
' places result in top left cell
' blanks all other cells in the range
Sub ConcatenateCells()
Attribute ConcatenateCells.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim dest As Range
    Dim resultStr As String
    Dim cellsToMerge As Range
    
    Set cellsToMerge = Application.Selection
    
    Set dest = cellsToMerge.Cells(1, 1)
    For Each cell In cellsToMerge
        resultStr = resultStr & " " & cell.Value
        cell.Value = ""
    Next
    
    resultStr = Trim(resultStr)
    dest.Value = resultStr
End Sub

' for each cell in the current selection, it removes any line breaks
Sub TextToSingleLine()
    Dim cellsToProcess As Range
'    Dim cell As Range
    
    Set cellsToProcess = Application.Selection

  cellsToProcess.Cells.Replace Chr(10), ";", xlPart
End Sub

' quick and easy way to use index() & match()
'
Function MatrixLU(MatrixRef As Range, RowRef, ColumnRef, ValueIfNotFound)

    If IsError(Application.Match(RowRef, MatrixRef.Columns(1), 0)) Then
        MatrixLU = ValueIfNotFound
        Exit Function
    End If
    
    If IsError(Application.Match(ColumnRef, MatrixRef.Rows(1), 0)) Then
        MatrixLU = ValueIfNotFound
        Exit Function
    End If
    
    If ErrorMsg = "" Then
        MatrixLU = MatrixRef(Application.Match(RowRef, MatrixRef.Columns(1), 0), _
        Application.Match(ColumnRef, MatrixRef.Rows(1), 0))
        Else
        MsgBox ErrorMsg
    End If
End Function

'returns the number of the last row in the sheet with data in it
'
Function lastRow(sh As Worksheet)
    Dim lastrownum As Long
    
    On Error Resume Next
    lastrownum = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    lastRow = lastrownum
    
    On Error GoTo 0
End Function

' returns the number of the lat column on the sheet with data in it
'
Function lastCol(sh As Worksheet)
    On Error Resume Next
    lastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

' extract the hyperlink from a given cell
Function getURL(rg As Range) As String
  Dim Hyper As Hyperlink
  Set Hyper = rg.Hyperlinks.Item(1)
  getURL = Hyper.Address
End Function

' returns the conversion rate into CAD for the given currency
' currencyName is the text in the first column of FOREXtable
' FOREXtable is a 2 column table converting the currencyName
' use different FOREXtables to convert into different currencies.
Function FOREX(currencyName As String, FOREXtable As Range) As Double
    
    FOREX = Application.WorksheetFunction.VLookup(currencyName, FOREXtable, 2, False)
    
End Function




