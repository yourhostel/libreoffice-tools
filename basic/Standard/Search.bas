REM  *****  BASIC  *****

' Search.bas

Sub ShowFormByLast()
    Dim oDialog         As Object  
    Dim oButtonSearch   As Object
    Dim oListenerSearch As Object
    Dim sResult         As String
    
    oDialog = CreateDlgSearch()
    
    ' === Кнопка "Пошук" ===
    oButtonSearch    = oDialog.getControl("SearchButton")
    
    ' === Обробник кнопки oButtonCancel ===
    oListenerSearch = CreateUnoListener("SearchButton_", "com.sun.star.awt.XActionListener")
    oButtonSearch.addActionListener(oListenerSearch)
    
    oDialog.execute()   

    ' === Очищення ===
    oButtonSearch.removeActionListener(oListenerSearch)
    oDialog.dispose() 
End Sub

Sub SearchButton_actionPerformed(oEvent As Object)
    Dim oDialog   As Object
    Dim oResSmpl  As Object
    Dim sLastSmpl As String
    
    oDialog   = oEvent.Source.getContext()
    sLastSmpl = oDialog.getControl("SearchField").getText()
    oResSmpl  = oDialog.getControl("ResultEdit")
    
    oResSmpl.setText(GetByLast(sLastSmpl))         
End Sub

Function CreateDlgSearch()    
    Dim oDialog      As Object
    Dim oDialogModel As Object
    
    oDialog      = CreateUnoService("com.sun.star.awt.UnoControlDialog")
    oDialogModel = CreateUnoService("com.sun.star.awt.UnoControlDialogModel")
      
    oDialog.setModel(oDialogModel)
    
    ' ==== Параметри діалогу ====    
    With oDialogModel
        .PositionX = 100
        .PositionY = 100
        .Width     = 250
        .Height    = 160
        .Title     = "Пошук"
    End With
    
    Dim gX As Long, gY As Long
    gX = 10 : gY = 15
    
    AddBackground(oDialogModel, BACKGROUND) 
       
    FieldTemplate(oDialogModel,    "Search", "Шукане прізвище:", gx, gY, "", 70, 100)
    
    AddEditTemplate(oDialogModel,  "Result", gx, 20 + gY, 230, 100, "Тут буде результат пошуку", True)

    AddButton(oDialogModel,  "SearchButton", "Пошук", 85 + gx, 125 + gY, 60, 14)
    
    oDialog.createPeer(CreateUnoService("com.sun.star.awt.ExtToolkit"), Null)
       
    CreateDlgSearch = oDialog
End Function

Function GetByLast(sLastSmpl As String) As String
    Dim oDoc      As Object : oDoc      = ThisComponent
    Dim oSheet    As Object : oSheet    = oDoc.Sheets.getByName("Data")
    Dim nRowCount As Long   : nRowCount = oSheet.getRows().getCount()
    Dim sIdName   As String : sIdName   = ""

    Dim sLast     As String : sLast     = "" ' прізвище
    Dim sPatr     As String : sPatr     = "" ' ім'я по батькові
    Dim sId       As String : sId       = "" ' id
    Dim sCOut     As String : sCOut     = "" ' виселення
    Dim en        As String : en = Chr(8194)
     
    For iRow = 3 To nRowCount - 1
        sId   = Trim(oSheet.getCellByPosition(19, iRow).String) ' T
        sLast = Trim(oSheet.getCellByPosition(1, iRow).String)  ' B

        If sId = "" Then Exit For
        
        If sLast = sLastSmpl Then
            sCOut = oSheet.getCellByPosition(4, iRow).String    ' E        
            sPatr = oSheet.getCellByPosition(2, iRow).String    ' C
                      
            If sIdName = "" Then
                sIdName = "виселення" & en & "|" & en & "id" & String(5, Chr(8194)) & "|" & en & _
                    "прізвище" & String(5, Chr(8194)) & "|" & en & "ім'я по батькові" & Chr(10) & _
                     String(30, Chr(8212)) & Chr(10)
            End If
                       
            sIdName = sIdName & FormatSearch(sCOut, sId, sLast, sPatr) & Chr(10)                      
        End If   
    Next iRow
    
    If sIdName = "" Then sIdName = sLastSmpl & " не знайдено"
    
    GetByLast = sIdName
End Function

Function FormatSearch(sCOut As String, _
                      sId   As String, _
                      sLast As String, _
                      sPatr As String) As String
                      
    Dim en As String : en = Chr(8194)
    Dim sOut As String

    ' Вирівнюємо поля
    If Len(sId) < 6 Then
        sId = sId & String(6 - Len(sId), en)
    End If

    If Len(sLast) < 12 Then
        sLast = sLast & String(12 - Len(sLast), en)
    End If

    sOut = sCOut & en & "|" & en & sId & en & "|" & en & sLast & en & "|" & en & sPatr
    FormatSearch = sOut
End Function
