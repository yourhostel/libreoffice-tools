REM  *****  BASIC  *****

Sub AddButton(oDialogModel As Object, _
			   Name As String, _
			   Label As String, _
               PositionX As Integer, _
               PositionY As Integer, _
               Width As Integer, _
               Height As Integer)

    Dim oButton As Object
    oButton = oDialogModel.createInstance("com.sun.star.awt.UnoControlButtonModel")
    oButton.Name = Name
    oButton.Label = Label
    oButton.PositionX = PositionX
    oButton.PositionY = PositionY
    oButton.Width = Width
    oButton.Height = Height
    oButton.PushButtonType = com.sun.star.awt.PushButtonType.OK

    oDialogModel.insertByName(Name, oButton)
End Sub