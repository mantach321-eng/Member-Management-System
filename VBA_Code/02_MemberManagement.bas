' ============================================
' إدارة العناصر
' Member Management Module
' ============================================

Option Explicit

' ============================================
' دالة إضافة عنصر جديد
' ============================================
Public Function AddMember(memberData As Object) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim ws As Worksheet
    
    Set ws = gDataSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' تعيين رقم تلقائي
    ws.Cells(lastRow, 1).Value = lastRow - 1
    ws.Cells(lastRow, 2).Value = memberData("الاسم")
    ws.Cells(lastRow, 3).Value = memberData("اسم_الأب")
    ws.Cells(lastRow, 4).Value = memberData("الشهرة")
    ws.Cells(lastRow, 5).Value = memberData("اسم_الأم")
    ws.Cells(lastRow, 6).Value = memberData("العمر")
    ws.Cells(lastRow, 7).Value = memberData("الفرقة")
    ws.Cells(lastRow, 8).Value = memberData("الفوج")
    ws.Cells(lastRow, 9).Value = memberData("رقم_الهاتف")
    ws.Cells(lastRow, 10).Value = memberData("عنوان_السكن")
    
    AutoSave
    AddMember = True
    
    Exit Function
ErrorHandler:
    MsgBox "خطأ في إضافة العنصر: " & Err.Description, vbCritical
    AddMember = False
End Function

' ============================================
' دالة تعديل عنصر
' ============================================
Public Function EditMember(rowIndex As Long, memberData As Object) As Boolean
    On Error GoTo ErrorHandler
    
    If gUserRole = 3 Then ' Read-Only
        MsgBox "ليس لديك صلاحية للتعديل", vbExclamation
        EditMember = False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = gDataSheet
    
    ws.Cells(rowIndex, 2).Value = memberData("الاسم")
    ws.Cells(rowIndex, 3).Value = memberData("اسم_الأب")
    ws.Cells(rowIndex, 4).Value = memberData("الشهرة")
    ws.Cells(rowIndex, 5).Value = memberData("اسم_الأم")
    ws.Cells(rowIndex, 6).Value = memberData("العمر")
    ws.Cells(rowIndex, 7).Value = memberData("الفرقة")
    ws.Cells(rowIndex, 8).Value = memberData("الفوج")
    ws.Cells(rowIndex, 9).Value = memberData("رقم_الهاتف")
    ws.Cells(rowIndex, 10).Value = memberData("عنوان_السكن")
    
    AutoSave
    EditMember = True
    
    Exit Function
ErrorHandler:
    MsgBox "خطأ في تعديل العنصر: " & Err.Description, vbCritical
    EditMember = False
End Function

' ============================================
' دالة حذف عنصر
' ============================================
Public Function DeleteMember(rowIndex As Long) As Boolean
    On Error GoTo ErrorHandler
    
    If gUserRole >= 3 Then ' Not Admin or Secretary
        MsgBox "ليس لديك صلاحية للحذف", vbExclamation
        DeleteMember = False
        Exit Function
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("هل أنت متأكد من رغبتك في حذف هذا العنصر؟", vbYesNo + vbQuestion, "تأكيد الحذف")
    
    If response <> vbYes Then
        DeleteMember = False
        Exit Function
    End If
    
    gDataSheet.Rows(rowIndex).Delete
    AutoSave
    DeleteMember = True
    
    Exit Function
ErrorHandler:
    MsgBox "خطأ في حذف العنصر: " & Err.Description, vbCritical
    DeleteMember = False
End Function

' ============================================
' دالة الحصول على عدد العناصر
' ============================================
Public Function GetTotalMembers() As Long
    Dim ws As Worksheet
    Set ws = gDataSheet
    GetTotalMembers = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
End Function
