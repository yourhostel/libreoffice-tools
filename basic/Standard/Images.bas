REM  *****  BASIC  *****

' Images.bas

' =====================================================
' === Процедура AddLogo ===============================
' =====================================================
' → Додає до діалогу картинку (логотип) за вказаними координатами та розмірами.
' → Підтягує файл з диска й вставляє у форму як ImageControl.
' → Масштабує зображення, щоб воно вмістилося у заданих розмірах.
Sub AddLogo(oDialogModel As Object, _
            sName        As String, _
            PositionX    As Integer, _
            PositionY    As Integer, _
            Width        As Integer, _
            Height       As Integer)

    Dim oImage As Object
    oImage = oDialogModel.createInstance("com.sun.star.awt.UnoControlImageControlModel")
    oImage.Name = sName
    oImage.PositionX = PositionX
    oImage.PositionY = PositionY
    oImage.Width = Width
    oImage.Height = Height
    oImage.ScaleImage = True
    oImage.ImageURL = ConvertToURL(PATH_TO_LOGO)
    oDialogModel.insertByName(sName, oImage)
End Sub

