VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private weekDayAbbrvs() As String
Private monthAbbrvs() As String

Private Sub Class_Initialize()
    ' use split() to assign the array, since VB doesn't have the syntax to do it any other way
    weekDayAbbrvs() = Split("Err,Sun,Mon,Tues,Wed,Thurs,Fri,Sat,Sun", ",")
    monthAbbrvs() = Split("Err,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec", ",")

End Sub
Private Sub Class_Terminate()
    Erase weekDayAbbrvs
    Erase monthAbbrvs
End Sub

Public Function getWorkDayAbbr(day As Integer, monthStr As String, year As Integer) As String
    Dim theDate As String
    Dim monthElement As Variant
 
    On Error GoTo exit_GetWorkDayAbbr
    
    If day < 1 Or IsEmpty(day) Then
        getWorkDayAbbr = ""
        Exit Function
    End If
    
    For Each monthElement In monthAbbrvs
        If monthStr = monthElement Then
            GoTo valid_month
        End If
    Next
    getWorkDayAbbr = "Invalid Month"
    Exit Function
valid_month:

    theDate = day & " " & monthStr & " " & year
    'MsgBox (theDate)
    'MsgBox (Weekday(theDate) & " = " & weekDayAbbrvs(Weekday(theDate)))
    
    getWorkDayAbbr = weekDayAbbrvs(Weekday(theDate))
    Exit Function
    
exit_GetWorkDayAbbr:
    getWorkDayAbbr = "invalid input"
        
End Function

Public Function getMonthNumFromAbbrv(monthStr As String) As Integer
    Dim i As Integer
    
    On Error GoTo exit_getMonthNumFromAbbrv
    
    For i = 1 To 12
        If monthAbbrvs(i) = monthStr Then
            getMonthNumFromAbbrv = i
            Exit Function
        End If
    Next
    
    getMonthNumFromAbbrv = -1
    Exit Function
    
exit_getMonthNumFromAbbrv:
    MsgBox ("Error #" & Err.Number & " - " & Err.Description)
End Function


