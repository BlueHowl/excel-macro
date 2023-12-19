Attribute VB_Name = "Module1"

Sub EnvoyerEmails()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Rng As Range
    Dim iRow As Long
    
    Dim WordPath As String
    WordPath = ThisWorkbook.Path & "\test.docx"

    ' Ouvrir App mail (outlook)
    Set OutApp = CreateObject("Outlook.Application")
    
    ' Ouvrir Microsoft Word
    Set WordApp = CreateObject("Word.Application")
    'WordApp.Visible = True
    

    ' Spécifiez la plage contenant les informations de la liste
    Set Rng = ThisWorkbook.Sheets("Feuille1").Range("A1:C4")

    ' Boucle à travers chaque ligne de la plage
    For iRow = 1 To Rng.Rows.Count
        ' Insérer les données dans le document Word
        
        ' Ouvrir le modèle Word
        Set WordDoc = WordApp.Documents.Open(WordPath)
        
        WordDoc.Bookmarks("Bookmark1").Range.Text = Rng.Cells(iRow, 2).Value
        WordDoc.Bookmarks("Bookmark2").Range.Text = Rng.Cells(iRow, 3).Value
        
        
        ' Fermer et enregistrer le document Word
        ' WordDoc.Save
        WordDoc.SaveAs ThisWorkbook.Path & "\documentTest" & iRow & ".docx"
        WordDoc.Close
    
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            ' Spécifiez l'adresse e-mail
            .To = Rng.Cells(iRow, 1).Value 'mail à la colonne 1 du excel
            ' Spécifiez le sujet, le corps et l'attachement
            .Subject = "Sujet de l'e-mail"
            .Body = "Corps de l'e-mail"
            .Attachments.Add ThisWorkbook.Path & "\documentTest" & iRow & ".docx"
            ' Envoyez l'e-mail
            .Send
        End With
        Set OutMail = Nothing
    Next iRow

    ' Fermer Microsoft Word
    WordApp.Quit

    Set WordApp = Nothing
    Set WordDoc = Nothing
    Set OutApp = Nothing
End Sub

