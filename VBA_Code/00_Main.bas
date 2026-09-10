' ============================================
' نظام إدارة العناصر
' Member Management System v1.0
' ============================================
' الملف الرئيسي للبرنامج
' Main Entry Point
' ============================================

Option Explicit

' متغيرات عامة
Public gCurrentUser As String
Public gUserRole As Integer ' 1=Admin, 2=Secretary, 3=Read-Only
Public gDataSheet As Worksheet
Public gBackupFolder As String

' ============================================
' دالة البداية
' ============================================
Public Sub InitializeSystem()
    On Error GoTo ErrorHandler
    
    Set gDataSheet = ThisWorkbook.Sheets("Data")
    gBackupFolder = ThisWorkbook.Path & "\Backups\"
    
    ' التحقق من وجود مجلد النسخ الاحتياطية
    If Dir(gBackupFolder, vbDirectory) = "" Then
        MkDir gBackupFolder
    End If
    
    ' عرض نموذج تسجيل الدخول
    frmLogin.Show vbModal
    
    ' إذا لم يقم المستخدم بتسجيل الدخول، إغلق البرنامج
    If gCurrentUser = "" Then
        ThisWorkbook.Close False
        Exit Sub
    End If
    
    ' عرض الشاشة الرئيسية
    frmMainDashboard.Show vbModeless
    
    Exit Sub
ErrorHandler:
    MsgBox "خطأ في تهيئة البرنامج: " & Err.Description, vbCritical
End Sub

' ============================================
' دالة إغلاق البرنامج
' ============================================
Public Sub CloseApplication()
    Unload frmMainDashboard
    ThisWorkbook.Close True
End Sub

' ============================================
' دالة فحص الصلاحيات
' ============================================
Public Function HasPermission(requiredRole As Integer) As Boolean
    HasPermission = (gUserRole <= requiredRole)
End Function

' ============================================
' دالة حفظ تلقائي
' ============================================
Public Sub AutoSave()
    On Error Resume Next
    ThisWorkbook.Save
End Sub
