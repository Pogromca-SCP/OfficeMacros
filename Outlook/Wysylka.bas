Attribute VB_Name = "Wysylka"
Public Sub Certyfikaty()
    ' Automatyczne wysy³anie certyfikatów
    Dim TemplateM As String
    Dim TemplateK As String
    Dim TemplatePass As String
    Dim Extras As String
    Dim Format As String
    
    ' Parametry poni¿ej mo¿na dowolnie zmieniaæ w ramach konfiguracji
    ' Do prawid³owego dzia³ania skrypt wymaga odpowiednich szablonów!
    
    ' Œcie¿ka do Twojego szablonu mêskiego. Jeœli p³eæ nie ma znaczenia mo¿na u¿yæ tego samego szablonu, co w TemplateK
    TemplateM = ""
    
    ' Œcie¿ka do Twojego szablonu damskiego. Jeœli p³eæ nie ma znaczenia mo¿na u¿yæ tego samego szablonu, co w TemplateM
    TemplateK = ""
    
    ' Œcie¿ka do twojego szablonu z has³em
    TemplatePass = ""
    
    ' Œcie¿ka do folderu z za³¹cznikami
    Extras = ""
    
    ' Format za³¹czników
    Format = ""
    
    ' Modyfikacja dalszej czêœci kodu tylko na w³asne ryzyko ;)
    
    Dim TargetMail As String
    Dim Target As String
    TargetMail = InputBox("Podaj e-mail docelowy:")
    Target = InputBox("Podaj Imiê i Nazwisko osoby docelowej:")
    
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
    ' Odwraca miejscami imiê i nazwisko
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
    ' Automatycznie tworzy wiadomoœæ opart¹ na podanym szablonie i z za³¹cznikiem
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
