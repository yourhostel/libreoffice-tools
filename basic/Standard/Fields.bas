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
                     NamePrefix   As String, _
                     LabelText    As String, _
                     PositionX    As Integer, _
                     PositionY    As Integer, _
                     vText        As String, _
                     WidthLabel   As Integer, _
                     WidthCombo   As Integer, _
                     ListOfPlaces As String)

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

' =====================================================
' === Процедура AddButton =============================
' =====================================================
' → Додає кнопку на діалогову форму.
' → Налаштовує її позицію, розміри, підпис та тип кнопки.
' → За замовчуванням кнопка Standard (не закриває діалог).
Sub AddButton(oDialogModel As Object, _ 
			  sName        As String, _
			  Label        As String, _
              PositionX    As Integer, _ 
              PositionY    As Integer, _
              Width        As Integer, _
              Height       As Integer, _
     Optional PushType     As Variant)
               
	Dim iPushType As Integer
	
    If IsMissing(PushType) Then
        iPushType = 0 ' За замовчуванням — Standard (не закриває вікно)
    Else
        iPushType = PushType
    End If
    
    Dim oButton As Object
    oButton = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oButton.Name = sName
    oButton.Label = Label
    oButton.PositionX = PositionX
    oButton.PositionY = PositionY
    oButton.Width = Width
    oButton.Height = Height
    oButton.PushButtonType = iPushType
    oDialogModel.insertByName(sName, oButton)
End Sub

' =====================================================
' === Процедура OptionGroupTemplate ===================
' =====================================================
' → Додає на діалог два радіокнопки з мітками над ними.
' → Параметри:
'   — NamePrefix: префікс для імен
'   — LabelLeft, LabelRight: тексти міток над кнопками
'   — PositionX, PositionY: координати лівої кнопки
'   — WidthEach: ширина кожної кнопки
'   — DefaultLeft: True, якщо ліва кнопка за замовчуванням вибрана
' =====================================================
Sub OptionGroupTemplate(oDialogModel As Object, _
                        NamePrefix   As String, _
                        LabelLeft    As String, _
                        LabelRight   As String, _
                        PositionX    As Integer, _
                        PositionY    As Integer, _
                        WidthEach    As Integer, _
               Optional DefaultLeft  As Boolean)

    If IsMissing(DefaultLeft) Then DefaultLeft = True

    ' ==== Ліва мітка ====
    Dim oLabelLeft As Object
    oLabelLeft = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabelLeft.Name = NamePrefix & "LeftLabel"
    oLabelLeft.Label = LabelLeft
    oLabelLeft.PositionX = PositionX
    oLabelLeft.PositionY = PositionY - 10
    oLabelLeft.Width = WidthEach
    oLabelLeft.Height = 10
    oDialogModel.insertByName(oLabelLeft.Name, oLabelLeft)

    ' ==== Права мітка ====
    Dim oLabelRight As Object
    oLabelRight = oDialogModel.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    oLabelRight.Name = NamePrefix & "RightLabel"
    oLabelRight.Label = LabelRight
    oLabelRight.PositionX = PositionX + WidthEach + 10
    oLabelRight.PositionY = PositionY - 10
    oLabelRight.Width = WidthEach
    oLabelRight.Height = 10
    oDialogModel.insertByName(oLabelRight.Name, oLabelRight)

    ' ==== Ліва радіокнопка ====
    Dim oOptionLeft As Object
    oOptionLeft = oDialogModel.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    oOptionLeft.Name = NamePrefix & "Left"
    oOptionLeft.PositionX = PositionX
    oOptionLeft.PositionY = PositionY
    oOptionLeft.Width = WidthEach
    oOptionLeft.Height = 12
    oOptionLeft.State = IIf(DefaultLeft, True, False)
    oDialogModel.insertByName(oOptionLeft.Name, oOptionLeft)

    ' ==== Права радіокнопка ====
    Dim oOptionRight As Object
    oOptionRight = oDialogModel.createInstance("com.sun.star.awt.UnoControlRadioButtonModel")
    oOptionRight.Name = NamePrefix & "Right"
    oOptionRight.PositionX = PositionX + WidthEach + 10
    oOptionRight.PositionY = PositionY
    oOptionRight.Width = WidthEach
    oOptionRight.Height = 12
    oOptionRight.State = IIf(DefaultLeft, False, True)
    oDialogModel.insertByName(oOptionRight.Name, oOptionRight)
End Sub

