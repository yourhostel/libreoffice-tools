REM  *****  BASIC  *****

' Utilities.bas

' ==== Службові функції, утиліти, структура Map ======
' =====================================================


' =====================================================
' === Функція Capitalize =============================
' =====================================================
' → Перетворює рядок у формат:
'   Перша літера велика, інші — маленькі.
' → Якщо рядок порожній — повертає порожній.
Function Capitalize(sText As String) As String
    If Len(sText) = 0 Then
        Capitalize = ""
    Else
        Capitalize = UCase(Left(sText, 1)) & LCase(Mid(sText, 2))
    End If
End Function

' =====================================================
' === Функція LockFields ==============================
' =====================================================
' → Блокує поля на формі (робить ReadOnly).
' → Параметр sFieldNames — перелік імен полів через крапку з комою.
' → Наприклад:
'    LockFields(oEvent, "DurationField;OffsetField")
Sub LockFields(oEvent As Object, sFieldNames As String)
	Dim FieldNames() As String
	FieldNames = Split(sFieldNames, ";")

    ' Отримуємо діалог з контексту події
    Dim oDialog As Object
    oDialog = oEvent.Source.getContext()

    ' Цикл по масиву імен полів
    Dim i As Integer
    For i = LBound(FieldNames) To UBound(FieldNames)
        oDialog.getControl(FieldNames(i)).Model.ReadOnly = True
    Next i

End Sub

' =====================================================
' ================ Структура Map ======================
' =====================================================
' ==== Простий аналог Map/Dictionary для LibreOffice ==
' ==== Реалізовано на базі масиву пар Key-Value =======
' =====================================================

' ==== Оголошення структури Key-Value =====
Type KeyValuePair
    Key As String
    Value As String
End Type

' =====================================================
' === Функція CreateMap ===============================
' =====================================================
' → Створює пусту мапу (масив пар Key-Value).
' → Повертає Variant-масив.
Function CreateMap() As Variant
    CreateMap = Array()
End Function

' =====================================================
' === Процедура MapPut ================================
' =====================================================
' → Додає пару (Key, Value) в мапу.
' → Якщо ключ існує — оновлює його значення.
' → Якщо ключа нема — додає новий.
Sub MapPut(ByRef Map As Variant, ByVal Key As String, ByVal Value As String)
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            Map(i).Value = Value
            Exit Sub
        End If
    Next i

    ' Додаємо новий ключ
    Dim NewPair As KeyValuePair
    NewPair.Key = Key
    NewPair.Value = Value

    If IsEmpty(Map) Then
        Map = Array(NewPair)
    Else
        Map = AppendArray(Map, NewPair)
    End If
End Sub

' =====================================================
' === Функція MapGet ==================================
' =====================================================
' → Повертає значення за ключем.
' → Якщо ключ не знайдено — повертає порожній рядок "".
Function MapGet(ByVal Map As Variant, ByVal Key As String) As String
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            MapGet = Map(i).Value
            Exit Function
        End If
    Next i
    MapGet = "" ' Якщо не знайдено
End Function

' =====================================================
' === Функція MapHasKey ===============================
' =====================================================
' → Перевіряє наявність ключа в мапі.
' → Повертає True — якщо ключ є.
' → Повертає False — якщо ключа нема.
Function MapHasKey(ByVal Map As Variant, ByVal Key As String) As Boolean
    Dim i As Integer
    For i = LBound(Map) To UBound(Map)
        If Map(i).Key = Key Then
            MapHasKey = True
            Exit Function
        End If
    Next i
    MapHasKey = False
End Function

' =====================================================
' === Функція AppendArray =============================
' =====================================================
' → Внутрішня допоміжна функція.
' → Додає новий елемент до масиву.
Function AppendArray(arr As Variant, item As Variant) As Variant
    Dim l As Long
    l = -1
    On Error Resume Next
    l = UBound(arr) + 1
    On Error GoTo 0

    If l = 0 And IsEmpty(arr) Then
        AppendArray = Array(item)
    Else
        Dim temp() As Variant
        ReDim temp(0 To l)
        Dim i As Integer
        For i = 0 To l - 1
            temp(i) = arr(i)
        Next i
        temp(l) = item
        AppendArray = temp
    End If
End Function