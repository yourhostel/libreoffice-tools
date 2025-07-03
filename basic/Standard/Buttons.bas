REM  *****  BASIC  *****

' Buttons.bas

' =====================================================
' === Процедура AddButton =============================
' =====================================================
' → Додає кнопку на діалогову форму.
' → Налаштовує її позицію, розміри, підпис та тип кнопки.
' → За замовчуванням кнопка Standard (не закриває діалог).
Sub AddButton(oDialogModel As Object, _
			   Name As String, _
			   Label As String, _
               PositionX As Integer, _
               PositionY As Integer, _
               Width As Integer, _
               Height As Integer, _
               Optional PushType As Variant)

	Dim iPushType As Integer

    If IsMissing(PushType) Then
        iPushType = 0 ' За замовчуванням — Standard (не закриває вікно)
    Else
        iPushType = PushType
    End If

    Dim oButton As Object
    oButton = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oButton.Name = Name
    oButton.Label = Label
    oButton.PositionX = PositionX
    oButton.PositionY = PositionY
    oButton.Width = Width
    oButton.Height = Height
    oButton.PushButtonType = iPushType
    oDialogModel.insertByName(Name, oButton)
End Sub
