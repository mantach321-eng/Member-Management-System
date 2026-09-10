' ============================================
' نموذج تسجيل الدخول
' Login Module
' ============================================

Option Explicit

' ============================================
' دالة التحقق من بيانات الدخول
' ============================================
Public Function ValidateLogin(username As String, password As String) As Boolean
    Dim correctPassword As String
    
    ' قاموس كلمات المرور (يجب تغييره في الإنتاج)
    Select Case LCase(username)
        Case "admin"
            correctPassword = "admin123"
        Case "secretary"
            correctPassword = "sec123"
        Case "user"
            correctPassword = "user123"
        Case Else
            ValidateLogin = False
            Exit Function
    End Select
    
    ValidateLogin = (password = correctPassword)
End Function

' ============================================
' دالة تعيين دور المستخدم
' ============================================
Public Function GetUserRole(username As String) As Integer
    Select Case LCase(username)
        Case "admin"
            GetUserRole = 1 ' Admin
        Case "secretary"
            GetUserRole = 2 ' Secretary
        Case "user"
            GetUserRole = 3 ' Read-Only
        Case Else
            GetUserRole = 0 ' No Access
    End Select
End Function
