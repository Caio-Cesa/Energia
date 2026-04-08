Attribute VB_Name = "Módulo2"
Sub apaga()
Range("A3:XFD1048576").ClearContents
End Sub
Sub Filtro()
linha_fim = Range("a1000000").End(xlUp).Row
Let dado = Range("a3:a" & linha_fim).Value 'fazer com range identifica quantidade necessaria automaticamente
    Sheets("Base").ListObjects("Valerio").Range.AutoFilter Field:=1, Criteria1:=Application.Transpose(dado), Operator:=xlFilterValues
End Sub
Sub Caso_Base()
Dim w1 As Worksheet
Dim w2 As Worksheet
Dim w3 As Worksheet
Dim w As String
Dim a As String
Set w1 = Sheets("Inicial")
If Range("b3").Value = "" Then
    MsgBox ("Sem dados no local correto")
    Exit Sub
Else
    Application.ScreenUpdating = False
    a = Range("b5").Value
    linha_fim = Range("b1000000").End(xlUp).Row
    linha = 7 + WorksheetFunction.Match("*RELATORIO COMPLETO DO SISTEMA*", Range("b:b"), 0)
    Range("b" & linha).Offset(0, 3).Select
    Cells(linha, 5).FormulaLocal = "=EXT.TEXTO($B" & linha & ";1;16)"
    Cells(linha, 6).FormulaLocal = "=EXT.TEXTO($B" & linha & ";16;9)"
    Cells(linha, 7).FormulaLocal = "=EXT.TEXTO($B" & linha & ";24;9)"
    Cells(linha, 8).FormulaLocal = "=EXT.TEXTO($B" & linha & ";32;9)"
    Cells(linha, 9).FormulaLocal = "=EXT.TEXTO($B" & linha & ";40;9)"
    Cells(linha, 10).FormulaLocal = "=EXT.TEXTO($B" & linha & ";48;13)"
    Cells(linha, 11).FormulaLocal = "=EXT.TEXTO($B" & linha & ";60;9)"
    Cells(linha, 12).FormulaLocal = "=EXT.TEXTO($B" & linha & ";68;9)"
    Cells(linha, 13).FormulaLocal = "=EXT.TEXTO($B" & linha & ";76;7)"
    Cells(linha, 14).FormulaLocal = "=EXT.TEXTO($B" & linha & ";82;14)"
    Cells(linha, 15).FormulaLocal = "=EXT.TEXTO($B" & linha & ";95;4)"
    Cells(linha, 16).FormulaLocal = "=EXT.TEXTO($B" & linha & ";98;9)"
    Cells(linha, 17).FormulaLocal = "=EXT.TEXTO($B" & linha & ";106;9)"
    Cells(linha, 18).FormulaLocal = "=EXT.TEXTO($B" & linha & ";114;9)"
    Cells(linha, 19).FormulaLocal = "=EXT.TEXTO($B" & linha & ";122;8)"
    Cells(linha, 20).FormulaLocal = "=EXT.TEXTO($B" & linha & ";129;7)"
    Cells(linha, 21).FormulaLocal = "=EXT.TEXTO($B" & linha & ";135;5)"
    Cells(linha, 22).FormulaLocal = "=EXT.TEXTO($B" & linha & ";139;10)"
    Cells(linha, 23).FormulaLocal = "=EXT.TEXTO($B" & linha & ";148;10)"
    Cells(linha, 24).FormulaLocal = "=EXT.TEXTO($B" & linha & ";157;7)"
    Cells(linha, 25).FormulaLocal = "=EXT.TEXTO($B" & linha & ";163;7)"
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Cells(linha_fim, 25)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Range("a1").Select
    ActiveWorkbook.Sheets.Add.Name = "Base"
    Set w2 = Sheets("Base")
    'fazer tratamento de erro caso ja tenha planilha com nome de base
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("a1").Select
    linha_fim = Range("a1048576").End(xlUp).Row
    Range("v4:v" & linha_fim).FormulaLocal = "=SEERRO(SES(A2=""  .............."";DESLOC(A2;2;0);A1="" X-------------X"";DESLOC(A1;3;0);CORRESP(""  .............."";A4:A30;0)>2;V3);""-"")"
    Range("w4:w" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";SE(V4=V3;DESLOC($J$1;LIN(W4);0);DESLOC($J$1;CORRESP(V4;A:A;0);0)))"
    Range("x4:x" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";(DESLOC($K$1;LIN(K4);0)))"
    Range("y4:y" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";DESLOC($C$1;LIN(C4);0))"
    Range("z4:z" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";DESLOC($F$1;LIN(F4);0))"
    Range("aa4:aa" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)=0;""-"";SE(V4=A4;B3;""-""))"
    Range("v4:aa" & linha_fim).Value = Range("v4:aa" & linha_fim).Value
    Range("v4:aa" & linha_fim).AutoFilter Field:=1, Criteria1:="-", Operator:=xlOr, Criteria2:="="
    Application.DisplayAlerts = False
    Range("v5:aa" & linha_fim).Delete
    Application.DisplayAlerts = True
    Columns("a:u").Delete
    Rows("1:2").Delete
    Range("a1").Value = "De"
    Range("b1").Value = "Para"
    Range("c1").Value = "Cir."
    Range("d1").Value = "Capacidade"
    Range("e1").Value = "Carregamento"
    Range("a3").AutoFilter
    linha_fim = Range("a1").End(xlDown).Row
    ActiveWorkbook.Sheets.Add.Name = "Tensao"
    Set w3 = Sheets("Tensao")
    w2.Range("a1:a" & linha_fim).Copy
    w3.Range("a1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    w2.Range("f1:f" & linha_fim).Copy
    w3.Range("b1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    w3.Range("b1").Value = "Tensão"
    Range("a1:b" & linha_fim).AutoFilter Field:=2, Criteria1:="-", Operator:=xlOr, Criteria2:="="
    Application.DisplayAlerts = False
    w3.Range("a2:b" & linha_fim).Delete
    Application.DisplayAlerts = True
    w3.Range("a3").AutoFilter
    w2.Select
    w2.Range("f:f").ClearContents
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$E$" & linha_fim), , xlYes).Name = "BASE"
    Range("c2:e" & linha_fim).Value = Range("c2:e" & linha_fim).Value
    col_fim = Range("a1").End(xlToRight).Column
    Cells(1, col_fim).NoteText a
    w3.Select
    linha_fim = Range("a1").End(xlDown).Row
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B$" & linha_fim), , xlYes).Name = "TENSAO"
    w3.Range("b2:b" & linha_fim).NumberFormat = "0.000"
    col_fim = Range("a1").End(xlToRight).Column
    Cells(1, col_fim).NoteText a
    w3.Range("a1").Select
    Application.ScreenUpdating = True
    w1.Select
End If
End Sub
Sub Ocorrencia()
Attribute Ocorrencia.VB_ProcData.VB_Invoke_Func = " \n14"
Dim w1 As Worksheet
Dim w2 As Worksheet
Dim w3 As Worksheet
Dim w4 As Worksheet
Dim w As String
Set w1 = Sheets("Inicial")
Set w2 = Sheets("Base")
Set w3 = Sheets("Tensao")
If Range("b3").Value = "" Then
    MsgBox ("Sem dados no local correto")
    Exit Sub
Else
    Application.ScreenUpdating = False
    a = Range("b5").Value
    linha_fim = Range("b1000000").End(xlUp).Row
    linha = 7 + WorksheetFunction.Match("*RELATORIO COMPLETO DO SISTEMA*", Range("b:b"), 0)
    Range("b" & linha).Offset(0, 3).Select
    Cells(linha, 5).FormulaLocal = "=EXT.TEXTO($B" & linha & ";1;16)"
    Cells(linha, 6).FormulaLocal = "=EXT.TEXTO($B" & linha & ";16;9)"
    Cells(linha, 7).FormulaLocal = "=EXT.TEXTO($B" & linha & ";24;9)"
    Cells(linha, 8).FormulaLocal = "=EXT.TEXTO($B" & linha & ";32;9)"
    Cells(linha, 9).FormulaLocal = "=EXT.TEXTO($B" & linha & ";40;9)"
    Cells(linha, 10).FormulaLocal = "=EXT.TEXTO($B" & linha & ";48;13)"
    Cells(linha, 11).FormulaLocal = "=EXT.TEXTO($B" & linha & ";60;9)"
    Cells(linha, 12).FormulaLocal = "=EXT.TEXTO($B" & linha & ";68;9)"
    Cells(linha, 13).FormulaLocal = "=EXT.TEXTO($B" & linha & ";76;7)"
    Cells(linha, 14).FormulaLocal = "=EXT.TEXTO($B" & linha & ";82;14)"
    Cells(linha, 15).FormulaLocal = "=EXT.TEXTO($B" & linha & ";95;4)"
    Cells(linha, 16).FormulaLocal = "=EXT.TEXTO($B" & linha & ";98;9)"
    Cells(linha, 17).FormulaLocal = "=EXT.TEXTO($B" & linha & ";106;9)"
    Cells(linha, 18).FormulaLocal = "=EXT.TEXTO($B" & linha & ";114;9)"
    Cells(linha, 19).FormulaLocal = "=EXT.TEXTO($B" & linha & ";122;8)"
    Cells(linha, 20).FormulaLocal = "=EXT.TEXTO($B" & linha & ";129;7)"
    Cells(linha, 21).FormulaLocal = "=EXT.TEXTO($B" & linha & ";135;5)"
    Cells(linha, 22).FormulaLocal = "=EXT.TEXTO($B" & linha & ";139;10)"
    Cells(linha, 23).FormulaLocal = "=EXT.TEXTO($B" & linha & ";148;10)"
    Cells(linha, 24).FormulaLocal = "=EXT.TEXTO($B" & linha & ";157;7)"
    Cells(linha, 25).FormulaLocal = "=EXT.TEXTO($B" & linha & ";163;7)"
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range(Selection, Cells(linha_fim, 25)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.Copy
    Range("a1").Select
    Sheets.Add
    w = ActiveSheet.Name
    Set w4 = Sheets(w)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.Range("a1").Select
    linha_fim = Range("a1048576").End(xlUp).Row
    Range("v4:v" & linha_fim).FormulaLocal = "=SEERRO(SES(A2=""  .............."";DESLOC(A2;2;0);A1="" X-------------X"";DESLOC(A1;3;0);CORRESP(""  .............."";A4:A30;0)>2;V3);""-"")"
    Range("w4:w" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";SE(V4=V3;DESLOC($J$1;LIN(W4);0);DESLOC($J$1;CORRESP(V4;A:A;0);0)))"
    Range("x4:x" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";(DESLOC($K$1;LIN(K4);0)))"
    Range("y4:y" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";DESLOC($C$1;LIN(C4);0))"
    Range("z4:z" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;""-"";DESLOC($F$1;LIN(F4);0))"
    Range("aa4:aa" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)=0;""-"";SE(V4=A4;B3;""-""))"
    Range("v4:aa" & linha_fim).Value = Range("v4:aa" & linha_fim).Value
    Range("v4:z" & linha_fim).AutoFilter Field:=1, Criteria1:="-", Operator:=xlOr, Criteria2:="="
    Application.DisplayAlerts = False
    Range("v5:z" & linha_fim).Delete
    Application.DisplayAlerts = True
    Columns("a:u").Delete
    Rows("1:2").Delete
    Range("a1").Value = "De"
    Range("b1").Value = "Para"
    Range("c1").Value = "Cir."
    Range("d1").Value = "Capacidade"
    Range("e1").Value = "Carregamento"
    Range("f1").Value = "Tensão"
    Range("A3").AutoFilter
    linha_fim = Range("a1").End(xlDown).Row
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$" & linha_fim), , xlYes).Name = "BASE"
    Range("c2:e" & linha_fim).Value = Range("c2:e" & linha_fim).Value
End If
'caso que são de relaçao ao caso base
w2.Select
linha_fim = Range("a1").End(xlDown).Row
col_fim = Range("a1").End(xlToRight).Column + 1
w2.Cells(1, col_fim).Select
ActiveCell.Value = "Caso " & col_fim - 5
w2.Cells(2, col_fim).FormulaLocal = "=SE([@De]&[@Para]=BASE_1[@De]&BASE_1[@Para];BASE_1[@Carregamento];"""")"
w2.Range(Cells(1, col_fim), Cells(linha_fim, col_fim)).Value = Range(Cells(1, col_fim), Cells(linha_fim, col_fim)).Value
w2.Range(Cells(2, col_fim), Cells(linha_fim, col_fim)).NumberFormat = "0.00%"
col_fim = Range("a1").End(xlToRight).Column
w2.Cells(1, col_fim).NoteText a
w3.Select
linha_fim = Range("a1").End(xlDown).Row
col_fim = Range("a1").End(xlToRight).Column + 1
w3.Cells(1, col_fim).Select
ActiveCell.Value = "Caso " & col_fim - 2
w3.Cells(2, col_fim).FormulaLocal = "=PROCV([@De];BASE_1;6;FALSO)"
w3.Range(Cells(1, col_fim), Cells(linha_fim, col_fim)).Value = Range(Cells(1, col_fim), Cells(linha_fim, col_fim)).Value
w3.Range(Cells(2, col_fim), Cells(linha_fim, col_fim)).NumberFormat = "0.000"
col_fim = Range("a1").End(xlToRight).Column
w3.Cells(1, col_fim).NoteText a
w4.Select
Application.DisplayAlerts = False
ActiveWindow.SelectedSheets.Delete
Application.DisplayAlerts = True
w1.Select
Application.ScreenUpdating = True
End Sub
