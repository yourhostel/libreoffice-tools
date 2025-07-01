REM  *****  BASIC  *****

' Notification.bas

Sub ShowDialog(Title As String, Message As String)

    Dim oDlg As Object, oDlgModel As Object
    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlg.setModel(oDlgModel)

    With oDlgModel
        .Title = Title
        .Width = 220
        .Height = 50
    End With

    ' ==== FieldTemplate тільки для мітки без поля ====
    ' NamePrefix
    ' LabelText
    ' PositionX
    ' PositionY
    ' vText
    ' WidthLabel
    ' WidthField
    ' ReadOnly
    Call FieldTemplate(oDlgModel, "Error", Message, 10, 20, "", 210, 0, True)

    ' ==== Кнопка OK ====
    Call AddButton(oDlgModel, "OKButton", "OK", 90, 30, 40, 14, 1)

    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDlg.execute()
    oDlg.dispose()

End Sub