Attribute VB_Name = "Module1"
Option Explicit

Sub SortIP_A2_LastRow()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ipRange As Range
    Dim dataArr() As Variant
    Dim i As Long, j As Long
    Dim tmpIP As Variant, tmpKey As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Set ipRange = ws.Range("A2:A" & lastRow)
    
    ReDim dataArr(1 To ipRange.Rows.Count, 1 To 2)
    
    ' 데이터 적재
    For i = 1 To ipRange.Rows.Count
        dataArr(i, 1) = ipRange.Cells(i, 1).Value
        dataArr(i, 2) = IPToLong(CStr(dataArr(i, 1))) ' ← CStr로 안전하게
    Next i
    
    ' 간단 정렬(버블; 행이 많으면 Range.Sort로 대체 권장)
    For i = 1 To UBound(dataArr, 1) - 1
        For j = i + 1 To UBound(dataArr, 1)
            If dataArr(i, 2) > dataArr(j, 2) Then
                tmpIP = dataArr(i, 1): dataArr(i, 1) = dataArr(j, 1): dataArr(j, 1) = tmpIP
                tmpKey = dataArr(i, 2): dataArr(i, 2) = dataArr(j, 2): dataArr(j, 2) = tmpKey
            End If
        Next j
    Next i
    
    ' 다시 쓰기
    For i = 1 To ipRange.Rows.Count
        ipRange.Cells(i, 1).Value = dataArr(i, 1)
    Next i
End Sub

Function IPToLong(ByVal ip As String) As Double
    Dim parts() As String
    Dim a As Long, b As Long, c As Long, d As Long
    
    ip = Trim(ip)
    If Len(ip) = 0 Then
        IPToLong = 9.9E+99   ' 공백/무효값은 맨 뒤로 보내기
        Exit Function
    End If
    
    parts = Split(ip, ".")
    If UBound(parts) <> 3 Then
        IPToLong = 9.9E+99   ' 무효한 IPv4 형식
        Exit Function
    End If
    
    a = Val(parts(0)): b = Val(parts(1)): c = Val(parts(2)): d = Val(parts(3))
    
    ' 0~255 범위 체크 (범위 밖이면 뒤로)
    If a < 0 Or a > 255 Or b < 0 Or b > 255 Or c < 0 Or c > 255 Or d < 0 Or d > 255 Then
        IPToLong = 9.9E+99
        Exit Function
    End If
    
    ' 16777216 = 256^3, 65536 = 256^2, 256 = 256^1
    IPToLong = CDbl(a) * 16777216# + CDbl(b) * 65536# + CDbl(c) * 256# + CDbl(d)
End Function

