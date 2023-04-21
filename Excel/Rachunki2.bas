Attribute VB_Name = "Rachunki"
Sub Rachunek()
Attribute Rachunek.VB_Description = "Generowanie nowego rachunku"
Attribute Rachunek.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' Rachunek Makro
' Generowanie nowego rachunku
'
' Klawisz skrótu: Ctrl+d
'

'
' Poni¿sze adresy komórek mo¿na dowolnie zmieniaæ:
' Komórka do wpisania numeru rachunku
    Const Num As String = "D2"
    
' Komórka do wpisania obecnej daty
    Const Dat As String = "F2"

' Komórka do wpisania pocz¹tku pracy
    Const Start As String = "B4"
    
' Komórka do wpisania zakoñczenia pracy
    Const Finish As String = "D4"
    
' Komórka do wpisania liczby godzin
    Const Hours As String = "F4"
    
' Komórka do wpisania kwoty
    Const Value As String = "F10"
    
' Komórka do kwoty s³ownej
    Const Spoken As String = "B25"
    
' Komórka do której mo¿na powróciæ po skoñczonej pracy
    Const Retur As String = "G1"
    
' Koniec parametrów, tego co jest dalej radzê nie zmieniaæ ;)

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
    ActiveCell.FormulaR1C1 = Left(ActiveCell.Value, 3) & Month & "/" & Year & "/R"
    Range(Start).Select
    ActiveCell.FormulaR1C1 = ProcessWorkPeriod(True)
    Range(Finish).Select
    ActiveCell.FormulaR1C1 = ProcessWorkPeriod(False)
    Range(Hours).Select
    ActiveCell.FormulaR1C1 = InputBox("Podaj liczbê godzin:", "Liczba godzin", ActiveCell.Value)
    Range(Value).Select
    ActiveCell.FormulaR1C1 = InputBox("Podaj kwotê:", "Kwota", ActiveCell.Value)
    Range(Spoken).Select
    ActiveCell.FormulaR1C1 = InputBox("Podaj kwotê (s³ownie):", "Kwota", ActiveCell.Value)
    Range(Retur).Select
End Sub

Private Function IntToText(Value As Integer) As String
    Dim ValStr As String
    ValStr = CStr(Value)
    IntToText = IIf(Value < 10, "0" & ValStr, ValStr)
End Function

Private Function ProcessWorkPeriod(Start As Boolean) As String
    Dim Today As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    Dim IsOdd As Boolean
    Today = Date
    M = Month(Today)
    Y = Year(Today)
    D = IIf(Start, 1, Day(Today))
    ProcessWorkPeriod = IntToText(D) & "." & IntToText(M) & "." & CStr(Y)
End Function
