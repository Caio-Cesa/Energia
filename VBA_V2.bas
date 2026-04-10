Attribute VB_Name = "Modulo_V2"

' ==============================================================================
' PROJETO VALERIO - VERSÃO 2.0 (FLUXO DUAS ETAPAS)
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
    
    caminhoPasta = ActiveWorkbook.Path & "\Dados_Entrada\"
    
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
        wTemp.Cells(linha, 2).Value = strLine ' Texto bruto na B
        linha = linha + 1
    Loop
    Close #iFile
    
    ' Localizar início do relatório
    On Error Resume Next
    start_match = 0
    start_match = Application.WorksheetFunction.Match("*RELATORIO COMPLETO DO SISTEMA*", wTemp.Range("B:B"), 0)
    On Error GoTo 0
    
    If start_match = 0 Then
        Application.DisplayAlerts = False
        wTemp.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    linha_fim = wTemp.Cells(wTemp.Rows.Count, 2).End(xlUp).Row
    
    ' ==========================================================================
    ' PASSO 1: EXTRAÇÃO BRUTA (EXT.TEXTO)
    ' ==========================================================================
    With wTemp
        ' Extraindo para as colunas E até Y (índices 5 a 25)
        .Range(.Cells(start_match, 5), .Cells(linha_fim, 5)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";1;16)"
        .Range(.Cells(start_match, 6), .Cells(linha_fim, 6)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";16;9)"
        .Range(.Cells(start_match, 7), .Cells(linha_fim, 7)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";24;9)"
        .Range(.Cells(start_match, 8), .Cells(linha_fim, 8)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";32;9)"
        .Range(.Cells(start_match, 9), .Cells(linha_fim, 9)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";40;9)"
        .Range(.Cells(start_match, 10), .Cells(linha_fim, 10)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";48;13)"
        .Range(.Cells(start_match, 11), .Cells(linha_fim, 11)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";60;9)"
        .Range(.Cells(start_match, 12), .Cells(linha_fim, 12)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";68;9)"
        .Range(.Cells(start_match, 13), .Cells(linha_fim, 13)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";76;7)"
        .Range(.Cells(start_match, 14), .Cells(linha_fim, 14)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";82;14)"
        .Range(.Cells(start_match, 15), .Cells(linha_fim, 15)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";95;4)"
        .Range(.Cells(start_match, 16), .Cells(linha_fim, 16)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";98;9)"
        .Range(.Cells(start_match, 17), .Cells(linha_fim, 17)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";106;9)"
        .Range(.Cells(start_match, 18), .Cells(linha_fim, 18)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";114;9)"
        .Range(.Cells(start_match, 19), .Cells(linha_fim, 19)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";122;8)"
        .Range(.Cells(start_match, 20), .Cells(linha_fim, 20)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";129;7)"
        .Range(.Cells(start_match, 21), .Cells(linha_fim, 21)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";135;5)"
        .Range(.Cells(start_match, 22), .Cells(linha_fim, 22)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";139;10)"
        .Range(.Cells(start_match, 23), .Cells(linha_fim, 23)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";148;10)"
        .Range(.Cells(start_match, 24), .Cells(linha_fim, 24)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";157;7)"
        .Range(.Cells(start_match, 25), .Cells(linha_fim, 25)).FormulaLocal = "=EXT.TEXTO($B" & start_match & ";163;7)"
        
        ' Converter extrações para Valores
        .Range("E" & start_match & ":Y" & linha_fim).Value = .Range("E" & start_match & ":Y" & linha_fim).Value
        
        ' =========================================================================
        ' PASSO 2: LÓGICA DE CATEGORIA (SES) SOBRE OS DADOS EXTRAÍDOS
        ' =========================================================================
        ' Coluna V (índice 22) baseada no que foi extraído para a Coluna E (índice 5)
        .Range("V" & start_match + 3 & ":V" & linha_fim).FormulaLocal = "=SEERRO(SES(E" & start_match + 1 & "=""  .............."";DESLOC(E" & start_match + 1 & ";2;0);E" & start_match & "="" X-------------X"";DESLOC(E" & start_match & ";3;0);CORRESP(""  .............."";E" & start_match + 3 & ":E" & start_match + 30 & ";0)>2;V" & start_match + 2 & ");""-"")"
        
        Calculate
        .Range("V" & start_match & ":V" & linha_fim).Value = .Range("V" & start_match & ":V" & linha_fim).Value
        
        ' Coluna Z (índice 26): Origem
        .Range("Z" & start_match & ":Z" & linha_fim).Value = nomeArquivo
        
        ' Filtrar lixo (Onde V = "-")
        .Range("V" & start_match & ":Z" & linha_fim).AutoFilter Field:=1, Criteria1:="=-", Operator:=xlOr, Criteria2:="="
        On Error Resume Next
        .Range("V" & start_match + 1 & ":V" & linha_fim).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        .AutoFilterMode = False
        
        ' =========================================================================
        ' PASSO 3: CONSOLIDAÇÃO
        ' =========================================================================
        lastRowBase = wBase.Cells(wBase.Rows.Count, 1).End(xlUp).Row
        If lastRowBase = 1 And wBase.Cells(1, 1).Value = "" Then lastRowBase = 0
        
        ' Copiar E:Y (Dados) e Z (Origem)
        .Range("E" & start_match & ":Y" & linha_fim).Copy
        wBase.Cells(lastRowBase + 1, 1).PasteSpecial xlPasteValues
        .Range("Z" & start_match & ":Z" & linha_fim).Copy
        wBase.Cells(lastRowBase + 1, 22).PasteSpecial xlPasteValues
    End With
    
    Application.DisplayAlerts = False
    wTemp.Delete
    Application.DisplayAlerts = True
End Sub

Sub FormatarComoTabela(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim tbl As ListObject
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 22
    
    ' Cabeçalhos Básicos
    ws.Cells(1, 1).Value = "De / Barra"
    ws.Cells(1, 2).Value = "Para"
    ws.Cells(1, 13).Value = "Carregamento"
    ws.Cells(1, 22).Value = "Origem_Caso"
    
    If lastRow > 1 Then
        On Error Resume Next
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        tbl.TableStyle = "TableStyleMedium2"
        ws.Columns("A:V").AutoFit
        On Error GoTo 0
    End If
End Sub
