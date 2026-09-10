' ============================================
' النسخ الاحتياطية
' Backup Module
' ============================================

Option Explicit

' ============================================
' دالة إنشاء نسخة احتياطية
' ============================================
Public Function CreateBackup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim backupPath As String
    Dim backupName As String
    Dim timeStamp As String
    
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    backupName = "Backup_" & timeStamp & ".xlsm"
    backupPath = gBackupFolder & backupName
    
    ThisWorkbook.SaveCopyAs backupPath
    
    MsgBox "تم حفظ النسخة الاحتياطية بنجاح" & vbCrLf & backupPath, vbInformation, "نسخة احتياطية"
    CreateBackup = True
    
    Exit Function
ErrorHandler:
    MsgBox "خطأ في إنشاء النسخة الاحتياطية: " & Err.Description, vbCritical
    CreateBackup = False
End Function

' ============================================
' دالة استعادة نسخة احتياطية
' ============================================
Public Function RestoreBackup() As Boolean
    On Error GoTo ErrorHandler
    
    Dim backupFile As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' البحث عن أحدث نسخة احتياطية
    backupFile = ""
    Dim file As Object
    Dim latestDate As Date
    latestDate = #1/1/1900#
    
    For Each file In FSO.GetFolder(gBackupFolder).Files
        If file.DateLastModified > latestDate Then
            latestDate = file.DateLastModified
            backupFile = file.Path
        End If
    Next file
    
    If backupFile = "" Then
        MsgBox "لم يتم العثور على نسخة احتياطية", vbExclamation
        RestoreBackup = False
        Exit Function
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("هل تريد استعادة النسخة الاحتياطية من: " & vbCrLf & backupFile & "?", vbYesNo + vbQuestion)
    
    If response <> vbYes Then
        RestoreBackup = False
        Exit Function
    End If
    
    ' نسخ النسخة الاحتياطية فوق الملف الحالي
    FSO.CopyFile backupFile, ThisWorkbook.Path & "\", True
    
    MsgBox "تمت استعادة النسخة الاحتياطية بنجاح. سيتم إعادة فتح الملف.", vbInformation
    ThisWorkbook.Close False
    
    RestoreBackup = True
    Exit Function
ErrorHandler:
    MsgBox "خطأ في استعادة النسخة الاحتياطية: " & Err.Description, vbCritical
    RestoreBackup = False
End Function

' ============================================
' دالة نسخ احتياطية تلقائية (يومية)
' ============================================
Public Sub AutomaticBackup()
    Dim lastBackupTime As String
    Dim settingsSheet As Worksheet
    
    ' إنشاء ورقة الإعدادات إذا لم تكن موجودة
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If settingsSheet Is Nothing Then
        Set settingsSheet = ThisWorkbook.Sheets.Add
        settingsSheet.Name = "Settings"
    End If
    
    lastBackupTime = settingsSheet.Range("A1").Value
    
    ' تحقق إذا مر يوم كامل
    If CDate(lastBackupTime) < Date - 1 Then
        CreateBackup
        settingsSheet.Range("A1").Value = Now
    End If
End Sub
