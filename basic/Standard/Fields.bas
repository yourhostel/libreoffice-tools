REM  *****  BASIC  *****

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
    oField.Height = 14
    oField.Text = vText
    oField.ReadOnly = bReadOnly
    oDialogModel.insertByName(oField.Name, oField)

End Sub
