' ============================================
' الإحصائيات والتقارير
' Statistics & Reports Module
' ============================================

Option Explicit

' ============================================
' دالة حساب إجمالي العناصر
' ============================================
Public Function GetTotalCount() As Long
    Dim ws As Worksheet
    Set ws = gDataSheet
    GetTotalCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
End Function

' ============================================
' دالة عد الأفواج
' ============================================
Public Function CountByBattalion(battalion As String) As Long
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim count As Long
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value = battalion Then
            count = count + 1
        End If
    Next i
    
    CountByBattalion = count
End Function

' ============================================
' دالة عد الفرق
' ============================================
Public Function CountByDivision(division As String) As Long
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim count As Long
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 7).Value = division Then
            count = count + 1
        End If
    Next i
    
    CountByDivision = count
End Function

' ============================================
' دالة عد العناصر بدون هاتف
' ============================================
Public Function CountNoPhone() As Long
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim count As Long
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Cells(i, 9).Value = "" Or ws.Cells(i, 9).Value = "0" Then
            count = count + 1
        End If
    Next i
    
    CountNoPhone = count
End Function

' ============================================
' دالة الحصول على الأعمار
' ============================================
Public Function GetAgeDistribution() As Object
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim ageData As Object
    Set ageData = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim lastRow As Long
    Dim age As Integer
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' تهيئة الفئات العمرية
    Dim ageRange As Variant
    ageRange = Array("10-15", "16-20", "21-25", "26-30", "31-40", "40+")
    Dim ageRangeKey As Variant
    For Each ageRangeKey In ageRange
        ageData(ageRangeKey) = 0
    Next ageRangeKey
    
    For i = 2 To lastRow
        age = Val(ws.Cells(i, 6).Value)
        
        If age <= 15 Then
            ageData("10-15") = ageData("10-15") + 1
        ElseIf age <= 20 Then
            ageData("16-20") = ageData("16-20") + 1
        ElseIf age <= 25 Then
            ageData("21-25") = ageData("21-25") + 1
        ElseIf age <= 30 Then
            ageData("26-30") = ageData("26-30") + 1
        ElseIf age <= 40 Then
            ageData("31-40") = ageData("31-40") + 1
        Else
            ageData("40+") = ageData("40+") + 1
        End If
    Next i
    
    Set GetAgeDistribution = ageData
End Function
