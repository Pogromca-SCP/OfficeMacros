Attribute VB_Name = "Maile"
Sub PobierzMaile()
Attribute PobierzMaile.VB_Description = "Pobiera i odznacza maile"
Attribute PobierzMaile.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' PobierzMaile Makro
' Pobiera i odznacza maile
'
' Klawisz skrótu: Ctrl+d
'
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim Num As Integer
    Dim D As Date
    Dim i As Integer
    
' Parametry do swobodnej modyfikacji:)
' Kolor jakim bêd¹ zaznaczane komórki(kolor sk³ada siê z trzech podstawowych wartoœci od 0 do 255: R - czerwony, G - zielony, B - niebieski)
    R = 255
    G = 255
    B = 0
' Liczba maili do pobrania
    Num = 30
' Koniec parametrów, tego co jest dalej radzê nie zmieniaæ;)

    If R < 0 Then R = 0 Else If R > 255 Then R = 255
    If G < 0 Then G = 0 Else If G > 255 Then G = 255
    If B < 0 Then B = 0 Else If B > 255 Then B = 255
    If Num < 0 Then Num = -Num Else If Num = 0 Then Exit Sub
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(R, G, B)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(0, 1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(R, G, B)
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    D = ActiveCell.Value
    i = 0
    Do While i < Num - 1
        ActiveCell.Offset(1, -1).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(R, G, B)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        ActiveCell.Offset(0, 1).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(R, G, B)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        ActiveCell.Value = D
        i = i + 1
    Loop
    ActiveCell.Offset(-(Num - 1), -1).Select
    Range(ActiveCell, ActiveCell.Offset(Num - 1, 0)).Select
    Selection.Copy
End Sub
