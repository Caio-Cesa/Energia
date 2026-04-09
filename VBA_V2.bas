Attribute VB_Name = "Modulo_V2"

' ==============================================================================
' PROJETO VALERIO - VERSÃO 2.0 (MODERNA & GLOBAL)
' ==============================================================================
' Melhorias:
' 1. Importação direta de arquivos .txt (sem necessidade de Ctrl+C / Ctrl+V)
' 2. Consolidação automática de múltiplos arquivos com identificação de origem.
' 3. Fórmulas globalizadas (Inglês) para compatibilidade internacional.
' 4. Otimização de performance e tratamento de erros.
' ==============================================================================

Sub Processar_Tudo()
    Dim fso As Object
    Dim pasta As Object
    Dim arquivo As Object
    Dim caminhoPasta As String
    Dim wBase As Worksheet, wTensao As Worksheet
    Dim contador As Long
    
    ' Configurações Iniciais
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    caminhoPasta = ActiveWorkbook.Path & "\Dados_Entrada\"
    
    ' Garantir que as abas existam e estejam limpas
    Set wBase = PrepararPlanilha("Base", True)
    Set wTensao = PrepararPlanilha("Tensao", True)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(caminhoPasta) Then
        MsgBox "Pasta 'Dados_Entrada' não encontrada no diretório: " & ActiveWorkbook.Path, vbCritical
        GoTo Finalizar
    End If
    
    Set pasta = fso.GetFolder(caminhoPasta)
    contador = 0
    
    ' Loop por cada arquivo .txt na pasta
    For Each arquivo In pasta.Files
        If Right(arquivo.Name, 4) = ".txt" Then
            contador = contador + 1
            Call Importar_E_Processar(arquivo.Path, arquivo.Name, wBase, wTensao, (contador = 1))
        End If
    Next arquivo
    
    If contador = 0 Then
        MsgBox "Nenhum arquivo .txt encontrado na pasta Dados_Entrada.", vbExclamation
    Else
        ' Finalizar formatação das tabelas
        FormatarComoTabela "Base", wBase
        FormatarComoTabela "Tensao", wTensao
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

Sub Importar_E_Processar(caminho As String, nomeArquivo As String, wBase As Worksheet, wTensao As Worksheet, ehPrimeiro As Boolean)
    Dim wTemp As Worksheet
    Dim iFile As Integer
    Dim strLine As String
    Dim linha As Long, linha_fim As Long
    Dim lastRowBase As Long, lastRowTensao As Long
    
    ' Criar aba temporária para o processamento do texto bruto
    Set wTemp = Sheets.Add
    iFile = FreeFile
    Open caminho For Input As #iFile
    
    linha = 1
    Do Until EOF(iFile)
        Line Input #iFile, strLine
        wTemp.Cells(linha, 2).Value = strLine ' Mantém na Coluna B para compatibilidade com lógica anterior
        linha = linha + 1
    Loop
    Close #iFile
    
    ' Localizar início do relatório
    On Error Resume Next
    linha = 0
    linha = Application.WorksheetFunction.Match("*RELATORIO COMPLETO DO SISTEMA*", wTemp.Range("B:B"), 0)
    On Error GoTo 0
    
    If linha = 0 Then
        Application.DisplayAlerts = False
        wTemp.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    linha_fim = wTemp.Cells(wTemp.Rows.Count, 2).End(xlUp).Row
    
    ' EXTRAÇÃO USANDO FÓRMULAS GLOBAIS (INGLÊS)
    ' Colunas V até AA como no original, mas traduzidas
    With wTemp
        ' Identificadores de categoria (Original V até AA)
        ' Coluna V: Identificador de Linha
        .Range("V" & linha & ":V" & linha_fim).Formula = "=IFERROR(IFS(B" & linha & "=""  .............."",OFFSET(B" & linha & ",2,0),B" & linha & "="" X-------------X"",OFFSET(B" & linha & ",3,0),TRUE,V" & linha - 1 & "),""-"")"
        ' Coluna W: "De"
        .Range("W" & linha & ":W" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",1,16))"
        ' Coluna X: "Para"
        .Range("X" & linha & ":X" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",16,9))"
        ' Coluna Y: "Cir"
        .Range("Y" & linha & ":Y" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",24,9))"
        ' Coluna Z: "Capacidade"
        .Range("Z" & linha & ":Z" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",32,9))"
        ' Coluna AA: "Carregamento"
        .Range("AA" & linha & ":AA" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",40,9))"
        ' Coluna AB: "Tensão"
        .Range("AB" & linha & ":AB" & linha_fim).Formula = "=IF(LEN(V" & linha & ")<2,""-"",MID(B" & linha & ",76,7))"
        ' Coluna AC: Origem (Nome do Arquivo)
        .Range("AC" & linha & ":AC" & linha_fim).Value = nomeArquivo
        
        ' Converter para Valores para agilizar
        .Range("V" & linha & ":AC" & linha_fim).Value = .Range("V" & linha & ":AC" & linha_fim).Value
        
        ' Filtrar e Limpar dados irrelevantes
        .Range("V" & linha & ":AC" & linha_fim).AutoFilter Field:=1, Criteria1:="=-", Operator:=xlOr, Criteria2:="="
        On Error Resume Next
        .Range("V" & linha + 1 & ":AC" & linha_fim).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        .AutoFilterMode = False
        
        ' Transferir para Base (Fluxos)
        lastRowBase = wBase.Cells(wBase.Rows.Count, 1).End(xlUp).Row
        If lastRowBase = 1 And wBase.Cells(1, 1).Value = "" Then lastRowBase = 0 ' Primeira execução
        
        ' Copiar W até AA e AC (De, Para, Cir, Cap, Carreg, Origem)
        .Range("W" & linha & ":AA" & linha_fim).Copy
        wBase.Cells(lastRowBase + 1, 1).PasteSpecial xlPasteValues
        .Range("AC" & linha & ":AC" & linha_fim).Copy
        wBase.Cells(lastRowBase + 1, 6).PasteSpecial xlPasteValues
        
        ' Transferir para Tensao
        lastRowTensao = wTensao.Cells(wTensao.Rows.Count, 1).End(xlUp).Row
        If lastRowTensao = 1 And wTensao.Cells(1, 1).Value = "" Then lastRowTensao = 0
        
        ' Copiar W (Nome da Barra), AB (Tensão) e AC (Origem)
        .Range("W" & linha & ":W" & linha_fim).Copy
        wTensao.Cells(lastRowTensao + 1, 1).PasteSpecial xlPasteValues
        .Range("AB" & linha & ":AB" & linha_fim).Copy
        wTensao.Cells(lastRowTensao + 1, 2).PasteSpecial xlPasteValues
        .Range("AC" & linha & ":AC" & linha_fim).Copy
        wTensao.Cells(lastRowTensao + 1, 3).PasteSpecial xlPasteValues
        
    End With
    
    ' Limpar aba temporária
    Application.DisplayAlerts = False
    wTemp.Delete
    Application.DisplayAlerts = True
End Sub

Sub FormatarComoTabela(topo As String, ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim tbl As ListObject
    
    ws.Select
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Adicionar Cabeçalhos se necessário
    If topo = "Base" Then
        ws.Cells(1, 1).Value = "De"
        ws.Cells(1, 2).Value = "Para"
        ws.Cells(1, 3).Value = "Circuito"
        ws.Cells(1, 4).Value = "Capacidade"
        ws.Cells(1, 5).Value = "Carregamento"
        ws.Cells(1, 6).Value = "Origem_Caso"
    Else
        ws.Cells(1, 1).Value = "Barra"
        ws.Cells(1, 2).Value = "Tensão"
        ws.Cells(1, 3).Value = "Origem_Caso"
    End If
    
    ' Criar Tabela
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    tbl.TableStyle = "TableStyleMedium2"
    
    ' Limpar linhas vazias ou de erro na Tensão (Filtro final)
    If topo = "Tensao" Then
        ws.Range("B:B").AutoFilter Field:=1, Criteria1:="<0.1", Operator:=xlOr, Criteria2:="-" ' Filtra lixo
        On Error Resume Next
        ws.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        ws.AutoFilterMode = False
    End If
End Sub
