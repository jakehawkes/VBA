Attribute VB_Name = "ChartAutomation"
Sub PT_ScaleChartAxis(SheetName As String, ChartName As String, X_or_Y As Variant, Primary_or_Secondary As Variant, _
    Minimum As Variant, Maximum As Variant, MajorUnit As Variant, MinorUnit As Variant)
  
  Dim wks As Worksheet, cht As Chart, ax As Axis
  Dim xyAxisGroup As XlAxisGroup
  Dim rCaller As Range
  Dim dMinimum As Double, dMaximum As Double
  Dim bSetMin As Boolean, bSetMax As Boolean
  Dim sError As String, iError As Long
  Dim vTestCategory As Variant
  
  DoEvents
  
  'Application.Volatile True
  
  If Len(SheetName) = 0 Then
    Set rCaller = Application.Caller ' cell containing UDF
    SheetName = rCaller.Parent.Name
  End If
  
  On Error Resume Next
  Set wks = Worksheets(SheetName)
  On Error GoTo 0
  If wks Is Nothing Then
    sError = "Error in Arguments to UpdateChart():" & vbCrLf & "Worksheet '" & SheetName & "' not found"
    GoTo ErrorFunction
  End If
  If wks.ChartObjects.Count = 0 Then
    sError = "Error in Arguments to UpdateChart():" & vbCrLf & "No charts found on worksheet '" & SheetName & "'"
    GoTo ErrorFunction
  End If
  
  If Len(ChartName) = 0 Then
    ChartName = wks.ChartObjects(1).Name
  End If
  
  On Error Resume Next
  Set cht = wks.ChartObjects(ChartName).Chart
  On Error GoTo 0
  If cht Is Nothing Then
    sError = "Error in Arguments to UpdateChart():" & vbCrLf & "Chart '" & ChartName & "' not found on worksheet '" & SheetName & "'"
    GoTo ErrorFunction
  End If
  
  Select Case LCase$(X_or_Y)
    Case "x", "1", "category", "cat"
      X_or_Y = xlCategory
      '' but not for non-value axes
    Case "y", "2", "value", "val"
      X_or_Y = xlValue
  End Select
  
  Select Case LCase$(Primary_or_Secondary)
    Case "primary", "pri", "1"
      Primary_or_Secondary = xlPrimary
    Case "secondary", "sec", "2"
      Primary_or_Secondary = xlSecondary
  End Select
  
  Set ax = cht.Axes(X_or_Y, Primary_or_Secondary)
  
  If ax.Type = xlCategory Then
    On Error Resume Next
    vTestCategory = ax.MinimumScale
    iError = Err.Number
    On Error GoTo 0
    If iError <> 0 Then
      sError = "Wrong Chart Type:" & vbCrLf & " Cannot scale a category-type axis"
      GoTo ErrorFunction
    End If
  End If
  
  If IsNumeric(Minimum) Or IsDate(Minimum) Then
    dMinimum = Minimum
    bSetMin = True
  Else
    Select Case LCase$(Minimum)
      Case "auto", "autoscale", "default"
        ax.MinimumScaleIsAuto = True
      Case "null", "skip", "ignore", "blank"
        ' make no change
      Case ""
        Minimum = "null"
        ' make no change
    End Select
  End If
  
  If IsNumeric(Maximum) Or IsDate(Maximum) Then
    dMaximum = Maximum
    bSetMax = True
  Else
    Select Case LCase$(Maximum)
      Case "auto", "autoscale", "default"
        ax.MaximumScaleIsAuto = True
      Case "null", "skip", "ignore", "blank"
        ' make no change
      Case ""
        Maximum = "null"
        ' make no change
    End Select
  End If
  
  If bSetMin And bSetMax Then
    If dMaximum <= dMinimum Then
      sError = "Error in Arguments to UpdateChart():" & vbCrLf & " Axis-maximum must be greater than Axis-minimum"
      GoTo ErrorFunction
    End If
  End If
  
  If bSetMin Then
    ax.MinimumScale = dMinimum
  End If
  
  If bSetMax Then
    ax.MaximumScale = dMaximum
  End If
  
  If IsNumeric(MajorUnit) Then
    If MajorUnit > 0 Then
      ax.MajorUnit = MajorUnit
    End If
  Else
    Select Case LCase$(MajorUnit)
      Case "auto", "autoscale", "default"
        ax.MajorUnitIsAuto = True
      Case "null", "skip", "ignore", "blank"
        ' make no change
      Case ""
        MajorUnit = "null"
        ' make no change
    End Select
  End If
  
  If IsNumeric(MinorUnit) Then
    If MinorUnit > 0 Then
      ax.MinorUnit = MinorUnit
    End If
  Else
    Select Case LCase$(MinorUnit)
      Case "auto", "autoscale", "default"
        ax.MinorUnitIsAuto = True
      Case "null", "skip", "ignore", "blank"
        ' make no change
      Case ""
        MinorUnit = "null"
        ' make no change
    End Select
  End If
  
  'PT_ScaleChartAxis = "Sheet '" & SheetName & "' Chart '" & ChartName & "' " _
      & Choose(Primary_or_Secondary, "Primary", "Secondary") & " " _
      & Choose(X_or_Y, "X", "Y") & " Axis " _
      & "{" & Minimum & ", " & Maximum & ", " & MajorUnit & ", " & MinorUnit & "}"
  
ExitFunction:
  Exit Sub
  
ErrorFunction:
  ' PT_ScaleChartAxis = sError
  Dim retVal As Variant
  
   retVal = MsgBox(sError, , "Update Chart Failed")
  GoTo ExitFunction
End Sub

Sub UpdateChart()
  Dim ChartMinAxis As Range, ChartMaxAxis As Range
  Set ChartMinAxis = ActiveSheet.Range("ChartMinAxis")
  Set ChartMaxAxis = ActiveSheet.Range("ChartMaxAxis")
  Call PT_ScaleChartAxis("WBS", "Chart 1", "Y", "Primary", ChartMinAxis, ChartMaxAxis, 7, 0)
End Sub
