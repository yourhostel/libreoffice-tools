REM  *****  BASIC  *****

' DeleteRecord.bas

' =====================================================
' === Процедура DeleteRow =============================
' =====================================================
' → Видаляє поточний рядок після запиту пароля
' → За бажанням — переносить курсор у стовпець A
Sub DeleteRow()
    Dim oDoc          As Object
    Dim oSheet        As Object
    Dim oSel          As Object
    Dim row           As Long
    Dim bWasProtected As Boolean

    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow
    oSheet = oDoc.Sheets(0)

    ' === Запитуємо пароль ===
    If Not ShowPasswordDialog(NEGET_RULES) Then
        MsgDlg "Відмова", String(18, " ") & "Операцію скасовано.", False, 50, 130
        Exit Sub
    End If
    
    ' === Перевіряємо, чи захищений аркуш ===
    bWasProtected = oSheet.IsProtected
    
    ' ==== Якщо захищений — знімаємо захист ====
    If bWasProtected Then
        oSheet.unprotect(NEGET_RULES)
    End If
    
    ' === Видаляємо рядок ===
    oSheet.Rows.removeByIndex(row, 1)
    
    ' ==== Повертаємо захист назад ====
    If bWasProtected Then
        oSheet.protect(NEGET_RULES)
    End If

    ' === Переносимо курсор у стовпець ===
    ' MoveCursorToColumnA()
End Sub
