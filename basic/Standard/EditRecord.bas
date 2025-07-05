REM  *****  BASIC  *****

' EditRecord.bas

Sub StartEdit()
    MsgBox "Start-1"
    If Not IsInAllowedColumns() Then Exit Sub
    MsgBox "Start"
    If Not IsInAllowedRows() Then Exit Sub

    MoveCursorToColumnA()
End Sub

Function IsInAllowedColumns() As Boolean
    Dim oDoc, oSel, col
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    col = oSel.RangeAddress.StartColumn

    If col < 0 Or col > 18 Then
        ShowDialog "Помилка", "Курсор має бути в межах стовпців A–S."
        IsInAllowedColumns = False
    Else
        IsInAllowedColumns = True
    End If
End Function

Function IsInAllowedRows() As Boolean
    Dim StartRow As Long, EndRow As Long
    Call FindEditRange(StartRow, EndRow)
    MsgBox "StartRow: " & StartRow & " EndRow: " & EndRow
    If StartRow = -1 Or EndRow = -1 Then
        ShowDialog "Помилка", "Не вдалося визначити діапазон редагування."
        IsInAllowedRows = False
        Exit Function
    End If

    Dim oDoc, oSel, row
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow

    ' якщо курсор поза межами дозволеного діапазону
    If row < StartRow Or row > EndRow Then
        ShowDialog "Помилка", "Курсор поза дозволеним діапазоном рядків: " & (StartRow+1) & "–" & (EndRow+1) & "."
        IsInAllowedRows = False
    Else
        IsInAllowedRows = True
    End If
End Function

Sub FindEditRange(ByRef StartRow As Long, ByRef EndRow As Long)
    Dim oDoc, oSheet, oRange, aData, i, rowCount
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)
    rowCount = oSheet.Rows.Count

    ' читаємо діапазон стовпця E починаючи з 4 рядка
    oRange = oSheet.getCellRangeByPosition(4, 3, 4, rowCount - 1)
    aData = oRange.getDataArray()

    StartRow = -1
    EndRow = -1

    ' йдемо вниз — шукаємо перший порожній
    For i = 0 To UBound(aData)
        If Trim(aData(i)(0)) = "" Then
            EndRow = i + 3 - 1 ' -1 бо порожня — вже за діапазоном
            Exit For
        End If
    Next i

    ' якщо не знайшли порожню — беремо самий низ
    If EndRow = -1 Then
        EndRow = rowCount - 1
    End If

    ' тепер йдемо вгору від EndRow — шукаємо "інкасація"
    For i = EndRow - 3 To 0 Step -1
        If LCase(Trim(aData(i)(0))) = "інкасація" Then
            StartRow = i + 3 + 1 ' +1 бо "інкасація" сама не входить
            Exit For
        End If
    Next i

    ' якщо не знайшли "інкасація" — помилка
    If StartRow = -1 Then
        MsgBox "Не знайдено 'інкасація' у стовпці E.", 48, "Помилка"
    End If

End Sub

Sub MoveCursorToColumnA()
    Dim oDoc, oCtrl, oSel, row
    oDoc = ThisComponent
    oSel = oDoc.CurrentSelection
    row = oSel.RangeAddress.StartRow

    oCtrl = oDoc.CurrentController
    oCtrl.select(oDoc.Sheets(0).getCellByPosition(0, row))
End Sub