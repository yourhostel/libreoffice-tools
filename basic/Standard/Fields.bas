REM  *****  BASIC  *****

' Fields.bas

' =====================================================
' === Процедура FieldTemplate =========================
' =====================================================
' → Додає на діалог мітку та поле введення (Edit) під нею.
' → Параметри:
'   — NamePrefix: префікс для імені (Label та Field)
'   — LabelText: текст мітки
'   — PositionX, PositionY: координати верхнього лівого кута поля
'   — vText: значення за замовчуванням
'   — WidthLabel: ширина мітки
'   — WidthField: ширина поля
'   — ReadOnly: необов’язково. True — тільки для читання.
Sub FieldTemplate(oDialogModel As Object, _
                  ByVal NamePrefix As String, _
                  ByVal LabelText As String, _
                  ByVal PositionX As Integer, _
                  ByVal PositionY As Integer, _
                  ByVal vText As String, _
                  ByVal WidthLabel As Integer, _
                  ByVal WidthField As Integer, _
                  Optional ByVal ReadOnly As Variant)

    Dim bReadOnly As Boolean
    If IsMissing(ReadOnly) Then
        bReadOnly = False
    Else
        bReadOnly = ReadOnly
    End If

    ' ==== Мітка ====
    Dim oLabel As Object
    oLabel = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabel.Name = NamePrefix & "Label"
    oLabel.Label = LabelText
    oLabel.PositionX = PositionX
    oLabel.PositionY = PositionY - 10 ' Мітка вище за поле
    oLabel.Width = WidthLabel
    oLabel.Height = 10
    oDialogModel.insertByName(oLabel.Name, oLabel)

    ' ==== Поле ====
    Dim oField As Object
    oField = oDialogModel.createInstance("com.sun.star.awt.UnoControlEditModel")
    oField.Name = NamePrefix & "Field"
    oField.PositionX = PositionX
    oField.PositionY = PositionY
    oField.Width = WidthField
    oField.Height = 15
    oField.Text = vText
    oField.ReadOnly = bReadOnly
    oDialogModel.insertByName(oField.Name, oField)

End Sub

' =====================================================
' === Процедура ComboBoxTemplate ======================
' =====================================================
' → Додає на діалог мітку та комбінований список (ComboBox) під нею.
' → Параметри:
'   — NamePrefix: префікс для імені (Label та Combo)
'   — LabelText: текст мітки
'   — PositionX, PositionY: координати верхнього лівого кута ComboBox
'   — vText: значення за замовчуванням
'   — WidthLabel: ширина мітки
'   — WidthCombo: ширина ComboBox
'   — ListOfPlaces: рядок зі значеннями через ;
Sub ComboBoxTemplate(oDialogModel As Object, _
                      ByVal NamePrefix As String, _
                      ByVal LabelText As String, _
                      ByVal PositionX As Integer, _
                      ByVal PositionY As Integer, _
                      ByVal vText As String, _
                      ByVal WidthLabel As Integer, _
                      ByVal WidthCombo As Integer, _
                      ByVal ListOfPlaces As String)

    ' ==== Мітка ====
    Dim oLabel As Object
    oLabel = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabel.Name = NamePrefix & "Label"
    oLabel.Label = LabelText
    oLabel.PositionX = PositionX
    oLabel.PositionY = PositionY - 10
    oLabel.Width = WidthLabel
    oLabel.Height = 10
    oDialogModel.insertByName(oLabel.Name, oLabel)

    ' ==== ComboBox ====
    Dim oCombo As Object
    oCombo = oDialogModel.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
    oCombo.Name = NamePrefix & "Combo"
    oCombo.PositionX = PositionX
    oCombo.PositionY = PositionY
    oCombo.Width = WidthCombo
    oCombo.Height = 15
    oCombo.Text = vText
    oCombo.StringItemList = Split(ListOfPlaces, ";")
    oDialogModel.insertByName(oCombo.Name, oCombo)
End Sub
