REM  *****  BASIC  *****

' Images.bas

' =====================================================
' === Процедура AddLogo ===============================
' =====================================================
' → Додає до діалогу картинку (логотип) за вказаними координатами та розмірами.
' → Підтягує файл з диска й вставляє у форму як ImageControl.
' → Масштабує зображення, щоб воно вмістилося у заданих розмірах.
Sub AddLogo(oDialogModel As Object, _
                  ByVal iName As String, _
                  ByVal PositionX As Integer, _
                  ByVal PositionY As Integer, _
             	  ByVal Width As Integer, _
             	  ByVal Height As Integer)

    Dim oImage As Object
    oImage = oDialogModel.createInstance("com.sun.star.awt.UnoControlImageControlModel")
    oImage.Name = iName
    oImage.PositionX = PositionX
    oImage.PositionY = PositionY
    oImage.Width = Width
    oImage.Height = Height
    oImage.ScaleImage = True
    oImage.ImageURL = ConvertToURL(PATH_TO_LOGO)
    oDialogModel.insertByName(iName, oImage)
End Sub
