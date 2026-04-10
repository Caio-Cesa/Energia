vba_code = """Attribute VB_Name = "Modulo_V2"

' ==============================================================================
' PROJETO VALERIO - VERSÃO 2.0 (EXTRAÇÃO COMPLETA - ANSI)
' ==============================================================================

Sub Processar_Tudo()
    Dim fso As Object
    Dim pasta As Object
    Dim arquivo As Object
    Dim caminhoPasta As String
    Dim wBase As Worksheet
    Dim contador As Long
    
    ' Configurações Iniciais
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    caminhoPasta = ActiveWorkbook.Path & "\\Dados_Entrada\\"
    
    ' Garantir que a aba mestre exista e esteja limpa
    Set wBase = PrepararPlanilha("Base", True)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(caminhoPasta) Then
        MsgBox "Pasta 'Dados_Entrada' não encontrada!", vbCritical
        GoTo Finalizar
    End If
    
    Set pasta = fso.GetFolder(caminhoPasta)
    contador = 0
    
    ' Loop por cada arquivo .txt na pasta
    For Each arquivo In pasta.Files
        If LCase(Right(arquivo.Name, 4)) = ".txt" Then
            contador = contador + 1
            Call Importar_E_Processar(arquivo.Path, arquivo.Name, wBase)
        End If
    Next arquivo
    
    If contador = 0 Then
        MsgBox "Nenhum arquivo .txt encontrado.", vbExclamation
    Else
        FormatarComoTabela wBase
        MsgBox "Processamento concluído! " & contador & " arquivos processados.", vbInformation
    End If

Finalizar:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Function PrepararPlanilha(nome As String, limpar As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = nome
    ElseIf limpar Then
        ws.Cells.Clear
    End If
    Set PrepararPlanilha = ws
End Function

Sub Importar_E_Processar(caminho As String, nomeArquivo As String, wBase As Worksheet)
    Dim wTemp As Worksheet
    Dim iFile As Integer
    Dim strLine As String
    Dim linha As Long, linha_fim As Long, start_match As Long
    Dim lastRowBase As Long
    
    Set wTemp = Sheets.Add
    iFile = FreeFile
    Open caminho For Input As #iFile
    
    linha = 1
    Do Until EOF(iFile)
        Line Input #iFile, strLine
        wTemp.Cells(linha, 1).Value = strLine
        linha = linha + 1
    Loop
    Close #iFile
    
    ' Localizar início do relatório
    On Error Resume Next
    start_match = 0
    start_match = Application.WorksheetFunction.Match("*RELATORIO COMPLETO DO SISTEMA*", wTemp.Range("A:A"), 0)
    On Error GoTo 0
    
    If start_match = 0 Then
        Application.DisplayAlerts = False
        wTemp.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    linha_fim = wTemp.Cells(wTemp.Rows.Count, 1).End(xlUp).Row
    
    ' EXTRAÇÃO COMPLETA COM FÓRMULAS LOCAIS (PORTUGUÊS)
    With wTemp
        ' Coluna V: Identificador de Linha (Categoria)
        .Range("V4:V" & linha_fim).FormulaLocal = "=SEERRO(SES(A2=\\"  ..............\\";DESLOC(A2;2;0);A1=\\" X-------------X\\";DESLOC(A1;3;0);CORRESP(\\"  ..............\\";A4:A30;0)>2;V3);\\"-\\")"
        
        ' Extrações Completas (20+ Campos conforme original)
        .Range("W4:W" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;1;16))"
        .Range("X4:X" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;16;9))"
        .Range("Y4:Y" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;24;9))"
        .Range("Z4:Z" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;32;9))"
        .Range("AA4:AA" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;40;9))"
        .Range("AB4:AB" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;48;13))"
        .Range("AC4:AC" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;60;9))"
        .Range("AD4:AD" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;68;9))"
        .Range("AE4:AE" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;76;7))"
        .Range("AF4:AF" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;82;14))"
        .Range("AG4:AG" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;95;4))"
        .Range("AH4:AH" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;98;9))"
        .Range("AI4:AI" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;106;9))"
        .Range("AJ4:AJ" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;114;9))"
        .Range("AK4:AK" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;122;8))"
        .Range("AL4:AL" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;129;7))"
        .Range("AM4:AM" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;135;5))"
        .Range("AN4:AN" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;139;10))"
        .Range("AO4:AO" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;148;10))"
        .Range("AP4:AP" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;157;7))"
        .Range("AQ4:AQ" & linha_fim).FormulaLocal = "=SE(NÚM.CARACT(V4)<2;\\"-\\";EXT.TEXTO(A4;163;7))"
        .Range("AR4:AR" & linha_fim).Value = nomeArquivo
        
        Calculate
        .Range("V4:AR" & linha_fim).Value = .Range("V4:AR" & linha_fim).Value
        
        ' Filtrar lixo
        .Range("V4:AR" & linha_fim).AutoFilter Field:=1, Criteria1:=\\"=-\\", Operator:=xlOr, Criteria2:=\\"=\\"
        On Error Resume Next
        .Range("V5:AR" & linha_fim).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        .AutoFilterMode = False
        
        lastRowBase = wBase.Cells(wBase.Rows.Count, 1).End(xlUp).Row
        If lastRowBase = 1 And wBase.Cells(1, 1).Value = "" Then lastRowBase = 0
        
        .Range("W4:AR" & linha_fim).Copy
        wBase.Cells(lastRowBase + 1, 1).PasteSpecial xlPasteValues
    End With
    
    Application.DisplayAlerts = False
    wTemp.Delete
    Application.DisplayAlerts = True
End Sub

Sub FormatarComoTabela(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim tbl As ListObject
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 22 ' Total de campos extraídos + Origem
    
    ' Cabeçalhos Básicos
    ws.Cells(1, 1).Value = "De / Barra"
    ws.Cells(1, 2).Value = "Para"
    ws.Cells(1, 3).Value = "Circuito"
    ws.Cells(1, 4).Value = "Capacidade"
    ws.Cells(1, 5).Value = "Carregamento"
    ws.Cells(1, 22).Value = "Origem_Caso"
    
    If lastRow > 1 Then
        On Error Resume Next
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        tbl.TableStyle = "TableStyleMedium2"
        ws.Columns("A:V").AutoFit
        On Error GoTo 0
    End If
End Sub
"""

with open(r"c:\\Users\\Usuario\\OneDrive\\Documentos\\Valerio\\VBA_V2.bas.tmp", "w", encoding="utf-8") as f:
    f.write(vba_code.replace("\\\\", "\\"))

# Convert to ANSI
import os
content = open(r"c:\\Users\\Usuario\\OneDrive\\Documentos\\Valerio\\VBA_V2.bas.tmp", "r", encoding="utf-8").read()
open(r"c:\\Users\\Usuario\\OneDrive\\Documentos\\Valerio\\VBA_V2.bas", "w", encoding="windows-1252").write(content)
"""
# Note: I'll use double backslashes for the paths inside the python code to avoid raw string issues
