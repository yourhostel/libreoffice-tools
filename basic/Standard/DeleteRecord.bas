REM  *****  BASIC  *****

' DeleteRecord.bas

' =====================================================
' === Процедура DeleteRow =============================
' =====================================================
' → Видаляє поточний рядок після запиту пароля
' → За бажанням — переносить курсор у стовпець A
Sub DeleteRow()
    Dim oDoc As Object, oSheet As Object, oSel As Object
    Dim row As Long

    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow
    oSheet = oDoc.Sheets(0)

    ' === Запитуємо пароль ===
    If Not ShowPasswordDialog(NEGET_RULES) Then
        MsgBox "Операцію скасовано.", 48, "Відмова"
        Exit Sub
    End If

    ' === Видаляємо рядок ===
    oSheet.Rows.removeByIndex(row, 1)

    ' === Переносимо курсор у стовпець ===
    ' MoveCursorToColumnA()
End Sub