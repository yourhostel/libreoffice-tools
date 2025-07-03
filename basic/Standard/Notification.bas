REM  *****  BASIC  *****

' Notification.bas

' =====================================================
' === Процедура ShowDialog ============================
' =====================================================
' → Створює невеликий діалог з заголовком та одним або двома повідомленнями.
' → Виводить повідомлення користувачу та чекає натискання кнопки OK.
' → Додаткове повідомлення (нижнє) є необов’язковим.
' → Автоматично підлаштовує висоту діалогу під кількість повідомлень.
Sub ShowDialog(Title As String, _
			   MessageTop As String, _
			   Optional MessageDown As Variant)

    Dim oDlg As Object, oDlgModel As Object
    Dim PosY_Button As Integer
    Dim Height_Dialog As Integer
    Dim CheckMessageDown As Boolean

    CheckMessageDown = IsMissing(MessageDown) Or Len(Trim(MessageDown)) = 0
    PosY_Button = 30
    Height_Dialog = 50

    If Not CheckMessageDown Then
    	PosY_Button = 40
    	Height_Dialog = 65
    End If

    oDlgModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
    oDlg = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDlg.setModel(oDlgModel)

    With oDlgModel
        .Title = Title
        .Width = 220
        .Height = Height_Dialog
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
    Call FieldTemplate(oDlgModel, "lineTop", MessageTop, 10, 20, "", 210, 0, True)

    If Not CheckMessageDown Then
        Call FieldTemplate(oDlgModel, "lineDown", MessageDown, 10, 35, "", 210, 0, True)
    	PosX_Button = 50
    End If

    ' ==== Кнопка OK ====
    Call AddButton(oDlgModel, "OKButton", "OK", 90, PosY_Button, 40, 14, 1)

    oDlg.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
    oDlg.execute()
    oDlg.dispose()
End Sub