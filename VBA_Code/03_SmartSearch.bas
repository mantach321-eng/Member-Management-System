' ============================================
' البحث الذكي
' Smart Search Module
' ============================================

Option Explicit

' ============================================
' دالة تطبيع الحروف العربية
' ============================================
Private Function NormalizeArabic(text As String) As String
    Dim result As String
    result = text
    
    ' استبدال الحروف المتشابهة
    result = Replace(result, "إ", "ا")
    result = Replace(result, "آ", "ا")
    result = Replace(result, "أ", "ا")
    result = Replace(result, "ى", "ي")
    result = Replace(result, "ة", "ه")
    
    NormalizeArabic = result
End Function

' ============================================
' دالة البحث الذكي
' ============================================
Public Function SmartSearch(searchQuery As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim results As Collection
    Set results = New Collection
    
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim searchWords() As String
    searchWords = Split(searchQuery, " ")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long, j As Long
    Dim found As Boolean
    Dim cellValue As String, normalizedCell As String
    Dim normalizedSearch As String
    
    For i = 2 To lastRow
        found = True
        
        ' البحث في كل كلمة من كلمات البحث
        For j = 0 To UBound(searchWords)
            normalizedSearch = NormalizeArabic(LCase(searchWords(j)))
            
            ' البحث في الأعمدة المحددة (2,3,4,8,9,8,7)
            Dim searchFields() As Integer
            searchFields = Array(1, 2, 3, 4, 8, 9, 7) ' الرقم والاسم واسم الأب والشهرة والفوج والفرقة
            
            Dim fieldFound As Boolean
            fieldFound = False
            
            Dim fieldIndex As Variant
            For Each fieldIndex In searchFields
                cellValue = ws.Cells(i, fieldIndex).Value & ""
                normalizedCell = NormalizeArabic(LCase(cellValue))
                
                If InStr(normalizedCell, normalizedSearch) > 0 Then
                    fieldFound = True
                    Exit For
                End If
            Next fieldIndex
            
            If Not fieldFound Then
                found = False
                Exit For
            End If
        Next j
        
        If found Then
            results.Add i ' إضافة رقم الصف
        End If
    Next i
    
    Set SmartSearch = results
    Exit Function
    
ErrorHandler:
    Set SmartSearch = New Collection
End Function

' ============================================
' دالة الحصول على بيانات العنصر
' ============================================
Public Function GetMemberData(rowIndex As Long) As Object
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    
    data("الرقم") = ws.Cells(rowIndex, 1).Value
    data("الاسم") = ws.Cells(rowIndex, 2).Value
    data("اسم_الأب") = ws.Cells(rowIndex, 3).Value
    data("الشهرة") = ws.Cells(rowIndex, 4).Value
    data("اسم_الأم") = ws.Cells(rowIndex, 5).Value
    data("العمر") = ws.Cells(rowIndex, 6).Value
    data("الفرقة") = ws.Cells(rowIndex, 7).Value
    data("الفوج") = ws.Cells(rowIndex, 8).Value
    data("رقم_الهاتف") = ws.Cells(rowIndex, 9).Value
    data("عنوان_السكن") = ws.Cells(rowIndex, 10).Value
    
    Set GetMemberData = data
End Function
