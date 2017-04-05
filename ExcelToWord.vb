 Sub ExcelToWord()
    
    Dim wrdApp      As Object
    Dim wrdDoc      As Object
    
    Dim Nblig       As Integer
    Dim NbCol       As Integer
    
    
    Dim Prompt      As String
    
    Dim PathTemplate As String
    Dim PathGeneratedFile As String
    
    
    PathTemplate = "C:\Users\Administrateur\Desktop\"
    PathGeneratedFile = "C:\Users\Administrateur\Desktop\FicheEvolution\"
    
    
            'Turn some stuff off while the macro is running
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    
            'count number of lines
    Nblig = Cells.SpecialCells(xlCellTypeLastCell).Row
    'MsgBox (Nblig)
    
    
            'count number of column
    Range("A1").Select
    NbCol = Selection.Column - 1
    Selection.CurrentRegion.Select
    Selection.Columns(Selection.Columns.Count).Select
     NbCol = Selection.Column - NbCol
    'MsgBox (NbCol)
    
    
    Set wrdApp = CreateObject("Word.Application")
   
        
     For i = 2 To Nblig
        MsgBox ("Génération de la Fiche Evolution N°" & i - 1)
        
            
                'Use the Status Bar to let the user know what the current progress is
        Prompt = "Copying Data: " & x - 1 & " of " & Nblig - 1 & "   (" & _
        Format((x - 1) / (Nblig - 1), "Percent") & ")"
        Application.StatusBar = Prompt
            
                'Open Template
         Set wrdDoc = wrdApp.Documents.Open(PathTemplate & "Doc.docx")
         wrdApp.Visible = True

                'If the file is not found, we need to end the sub and let the user know
        If wrdDoc Is Nothing Then
            MsgBox "Unable to find the Word file.", vbCritical, "File Not Found"
            wrdApp.Quit
            Set wrdApp = Nothing
            Exit Sub
        End If
        
            
                'Replace all bookmarks
        wrdDoc.Bookmarks("Nom_FE").Range.Text = Cells(i, 1)
        wrdDoc.Bookmarks("Num_Devis").Range.Text = Cells(i, 2)
        wrdDoc.Bookmarks("Nom_Projet").Range.Text = Cells(i, 3)
        wrdDoc.Bookmarks("Nom_DP_Client").Range.Text = Cells(i, 4)
        wrdDoc.Bookmarks("Num_DP_client").Range.Text = Cells(i, 5)
        wrdDoc.Bookmarks("Code_PAI").Range.Text = Cells(i, 6)
        wrdDoc.Bookmarks("Nom_CP_SQLI").Range.Text = Cells(i, 7)
        wrdDoc.Bookmarks("Num_CP_SQLI").Range.Text = Cells(i, 8)
        wrdDoc.Bookmarks("Date_envoie").Range.Text = Cells(i, 9)
        wrdDoc.Bookmarks("Description").Range.Text = Cells(i, 10)
        wrdDoc.Bookmarks("info_livrable_jalons").Range.Text = Cells(i, 11)
        wrdDoc.Bookmarks("NB_jour").Range.Text = Cells(i, 12)
        wrdDoc.Bookmarks("TJM").Range.Text = Cells(i, 13)
        
        wrdDoc.Bookmarks("SignetDate").Range.Text = Format(Now, "dd/mm/yyyy")
        
                 'Save file
        wrdDoc.SaveAs (PathGeneratedFile & "Result" & i & ".docx")
        
    
    
    Next
    

            'Close Template
    wrdDoc.Close
    wrdApp.Quit
    
    End Sub
    
    
    
    




