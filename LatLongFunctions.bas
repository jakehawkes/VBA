Attribute VB_Name = "LatLongFunctions"
' Latitude, Longitude, And Great Circle Distances
' By Chip Pearson
'    www.cpearson.com/Excel/LatLong.aspx
'    chip@ cpearson.com
'    17-Jan-2009

Option Explicit


Private Const C_RADIUS_EARTH_KM As Double = 6371.1
Private Const C_RADIUS_EARTH_MI As Double = 3958.82
Private Const C_PI As Double = 3.14159265358979

Function GreatCircleDistance(Latitude1 As Double, Longitude1 As Double, _
            Latitude2 As Double, Longitude2 As Double, _
            ValuesAsDecimalDegrees As Boolean, _
            ResultAsMiles As Boolean) As Double

Dim Lat1 As Double
Dim Lat2 As Double
Dim Long1 As Double
Dim Long2 As Double
Dim X As Long
Dim Delta As Double

If ValuesAsDecimalDegrees = True Then
    X = 1
Else
    X = 24
End If

' convert to decimal degrees
Lat1 = Latitude1 * X
Long1 = Longitude1 * X
Lat2 = Latitude2 * X
Long2 = Longitude2 * X

' convert to radians: radians = (degrees/180) * PI
Lat1 = (Lat1 / 180) * C_PI
Lat2 = (Lat2 / 180) * C_PI
Long1 = (Long1 / 180) * C_PI
Long2 = (Long2 / 180) * C_PI

' get the central spherical angle
Delta = ((2 * ArcSin(Sqr((Sin((Lat1 - Lat2) / 2) ^ 2) + _
    Cos(Lat1) * Cos(Lat2) * (Sin((Long1 - Long2) / 2) ^ 2)))))
    
If ResultAsMiles = True Then
    GreatCircleDistance = Delta * C_RADIUS_EARTH_MI
Else
    GreatCircleDistance = Delta * C_RADIUS_EARTH_KM
End If

End Function

Function ArcSin(X As Double) As Double
    ' VBA doesn't have an ArcSin function. Improvise
    ArcSin = Atn(X / Sqr(-X * X + 1))
End Function

Sub findNearestStation()
    Dim stnRange As Range
    Dim camRange As ListObject
    Dim stnWS As Worksheet
    Dim camWS As Worksheet
    Dim stnRow As Range
    Dim stnLat As Double
    Dim stnLong As Double
    Dim camLat As Double
    Dim camLong As Double
    Dim distance As Double
    
    Set stnWS = Worksheets("WSO Stations")
    Set camWS = Worksheets("Cameras")
    
    ' ignore header row, and find the range of stations
    With stnWS
        Set stnRange = .Range("A2", Cells(2, 1).End(xlDown).End(xlToRight))
    End With
    
    ' the camera table came from the XML, so we can refer to it by name
    Set camRange = camWS.ListObjects("Table1")
    'camRange.Select
    
    ' for each station
    For Each stnRow In stnRange.Rows
        stnLat = stnRow.Cells(1, 2)
        stnLong = stnRow.Cells(1, 3)
        
        ' for each camera
        For Each camRow In camRange.ListRows
            camLat = camRow.Range.Cells(1, camRange.ListColumns("Latitude").Index)
            camLong = camRow.Range.Cells(1, camRange.ListColumns("Longitude").Index)
            camNumber = camRow.Range.Cells(1, camRange.ListColumns("Number").Index)
            ' find the distance
            distance = GreatCircleDistance(stnLat, stnLong, camLat, camLong, True, False)
            If distance < 1 Then
                ' MsgBox ("Cam #" & camNumber & " is " & distance & " km away")
                stnRow.Cells(1, 9).Value = "Cam " & camNumber
                
            End If
            
        Next camRow
    Next stnRow
End Sub
