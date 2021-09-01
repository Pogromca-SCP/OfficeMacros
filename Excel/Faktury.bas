Attribute VB_Name = "Faktury"
Sub NowaFaktura()
Attribute NowaFaktura.VB_Description = "Generowanie nowej faktury"
Attribute NowaFaktura.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' NowaFaktura Makro
' Generowanie nowej faktury
'

'
' Poni¿sze adresy komórek mo¿na dowolnie zmieniaæ:
' Komórka do wpisania obecnej daty
    Const Dat As String = "H4:I4"
    
' Komórka do wpisania numeru faktury
    Const Num As String = "C8"
    
' Komórka do wpisania kwoty
    Const Val As String = "F19"
    
' Komórka do wpisania s³ownej kwoty
    Const Text As String = "B30"
    
' Komórka do której mo¿na powróciæ po skoñczonej pracy
    Const Ret As String = "A1"
    
' Koniec parametrów, tego co jest dalej radze nie zmieniaæ ;)

    Range(Dat).Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Dim Month As String
    Dim Year As String
    Month = Right(Left(ActiveCell.Value, 5), 2)
    Year = Right(ActiveCell.Value, 4)
    Range(Num).Select
    ActiveCell.FormulaR1C1 = Left(ActiveCell.Value, 3) & Month & "/" & Year
    Range(Val).Select
    ActiveCell.FormulaR1C1 = InputBox("Podaj kwotê (liczby):", "Kwota", ActiveCell.Value)
    Range(Text).Select
    ActiveCell.FormulaR1C1 = InputBox("Podaj kwotê (s³ownie):", "Kwota", ActiveCell.Value)
    Range(Ret).Select
End Sub
