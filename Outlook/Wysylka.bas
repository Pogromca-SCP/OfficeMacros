Attribute VB_Name = "Wysylka"
Public Sub Certyfikaty()
    ' Automatyczne wysy�anie certyfikat�w
    Dim TemplateM As String
    Dim TemplateK As String
    Dim TemplatePass As String
    Dim Extras As String
    Dim Format As String
    
    ' Parametry poni�ej mo�na dowolnie zmienia� w ramach konfiguracji
    ' Do prawid�owego dzia�ania skrypt wymaga odpowiednich szablon�w!
    
    ' �cie�ka do Twojego szablonu m�skiego. Je�li p�e� nie ma znaczenia mo�na u�y� tego samego szablonu, co w TemplateK
    TemplateM = ""
    
    ' �cie�ka do Twojego szablonu damskiego. Je�li p�e� nie ma znaczenia mo�na u�y� tego samego szablonu, co w TemplateM
    TemplateK = ""
    
    ' �cie�ka do twojego szablonu z has�em
    TemplatePass = ""
    
    ' �cie�ka do folderu z za��cznikami
    Extras = ""
    
    ' Format za��cznik�w
    Format = ""
    
    ' Modyfikacja dalszej cz�ci kodu tylko na w�asne ryzyko ;)
    
    Dim TargetMail As String
    Dim Target As String
    TargetMail = InputBox("Podaj e-mail docelowy:")
    Target = InputBox("Podaj Imi� i Nazwisko osoby docelowej:")
    
    If TargetMail = "" Or Target = "" Then
        MsgBox "Nie podano maila lub/oraz imienia i nazwiska!"
    Else
        IsMale.Show
        Male = IsMale.CheckBox.Value
        
        If Male Then
            Create TemplateM, TargetMail, Target, Extras, Format
        Else
            Create TemplateK, TargetMail, Target, Extras, Format
        End If
        
        Create TemplatePass, TargetMail, Target, "", ""
    End If
End Sub
Public Function Rev(Str As String) As String
    ' Odwraca miejscami imi� i nazwisko
    SplitId = InStrRev(Str, " ")
    If SplitId > 0 Then
        S1 = Left(Str, SplitId - 1)
        S2 = Right(Str, Len(Str) - SplitId)
        Rev = S2 & " " & S1
    Else
        Rev = Str
    End If
End Function
Public Sub Create(TempPath As String, TargetMail As String, Target As String, Extras As String, Format As String)
    ' Automatycznie tworzy wiadomo�� opart� na podanym szablonie i z za��cznikiem
    If Not TempPath = "" Then
        DoAttach = Not (Extras = "" Or Format = "")
        Dim Mail As MailItem
        Set Mail = Outlook.Application.CreateItemFromTemplate(TempPath)
        With Mail
            .To = TargetMail
            .Subject = Mail.Subject & Target
            
            If DoAttach Then
                .Attachments.Add (Extras & "\" & Rev(Target) & "." & Format)
            End If
        End With
        Mail.Display
        Set Mail = Nothing
    End If
End Sub
