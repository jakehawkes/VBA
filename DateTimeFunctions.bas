Attribute VB_Name = "DateTimeFunctions"
Public MyDate As New DateTimeClass

Function WeekDayAbbrFromDate(day As Integer, month As String, year As Integer)
    WeekDayAbbrFromDate = MyDate.getWorkDayAbbr(day, month, year)
End Function

Function MonthNumFromAbbrv(monthStr As String) As Integer
    Dim ret As Integer
    ret = MyDate.getMonthNumFromAbbrv(monthStr)
    MonthNumFromAbbrv = ret
End Function

Function isSDO(year As Integer, month As String, day As Integer, SDOrange As Range) As Boolean
    Dim theDate As Variant
  
    On Error Resume Next ' GoTo exit_isSDO
    
    theDate = DateSerial(year, MonthNumFromAbbrv(month), day)
    
    For Each cell In SDOrange
        If cell.Value = theDate Then
            isSDO = True
            Exit Function
        End If
    Next

    isSDO = False
End Function

Function WhatHoliday(year As Integer, month As String, day As Integer, holidayRange As Range) As String
    Dim theDate As Variant
    Dim row As Integer
    
    On Error Resume Next ' GoTo exit_isSDO
    
    theDate = DateSerial(year, MonthNumFromAbbrv(month), day)
    
    For row = 1 To holidayRange.Rows.Count
        If holidayRange.Cells(row, 1).Value = theDate Then
            WhatHoliday = holidayRange.Cells(row, 2).Value
            Exit Function
        End If
    Next

    WhatHoliday = ""
 
End Function
