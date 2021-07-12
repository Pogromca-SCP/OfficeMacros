Attribute VB_Name = "Poczta"
Sub Wpiszpriorytetpolecony()
Attribute Wpiszpriorytetpolecony.VB_Description = "Wpisuje automatycznie ""priorytet polecony"""
Attribute Wpiszpriorytetpolecony.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Wpiszpriorytetpolecony Makro
' Wpisuje automatycznie "priorytet polecony"
'
' Klawisz skrótu: Ctrl+p
'
    ActiveCell.FormulaR1C1 = "priorytet polecony"
End Sub
Sub Wpiszpriorytetzwykly()
Attribute Wpiszpriorytetzwykly.VB_Description = "Automatycznie wpisuje ""priorytet"""
Attribute Wpiszpriorytetzwykly.VB_ProcData.VB_Invoke_Func = "k\n14"
'
' Wpiszpriorytetzwykly Makro
' Automatycznie wpisuje "priorytet"
'
' Klawisz skrótu: Ctrl+k
'
    ActiveCell.FormulaR1C1 = "priorytet"
End Sub
Sub Polecony()
Attribute Polecony.VB_Description = "Automatycznie wpisuje ""polecony"""
Attribute Polecony.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Polecony Makro
' Automatycznie wpisuje "polecony"
'
' Klawisz skrótu: Ctrl+i
'
    ActiveCell.FormulaR1C1 = "polecony"
End Sub
Sub Wpiszdate()
Attribute Wpiszdate.VB_Description = "Autmatycznie wpisuje dzisiejsz¹ datê"
Attribute Wpiszdate.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' Wpiszdate Makro
' Automatycznie wpisuje dzisiejsz¹ datê
'
' Klawisz skrótu: Ctrl+d
'
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
