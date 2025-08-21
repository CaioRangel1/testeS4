' === CONEXÃO COM SAP ===
If Not IsObject(application) Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

' === CONEXÃO E CRIAÇÃO DO EXCEL ===
Dim objExcel, objWorkbook, objSheet, caminhoArquivo, fso

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

caminhoArquivo = "C:\Temp\cenario.xlsx"

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists("C:\Temp") Then
    fso.CreateFolder("C:\Temp")
End If

' Verifica se o arquivo Excel já existe
If fso.FileExists(caminhoArquivo) Then
    ' Se existe, abre o arquivo
    Set objWorkbook = objExcel.Workbooks.Open(caminhoArquivo)
    On Error Resume Next
    Set objSheet = objWorkbook.Sheets("Tabela Contratos")
    If objSheet Is Nothing Then
        Set objSheet = objWorkbook.Sheets.Add(objWorkbook.Sheets(1))
        objSheet.Name = "Tabela Contratos"
    End If
    On Error GoTo 0
Else
    ' Se não existe, cria um novo workbook
    Set objWorkbook = objExcel.Workbooks.Add
    Set objSheet = objWorkbook.Sheets(1)
    objSheet.Name = "Tabela Contratos"
    
    ' --- LINHA CRÍTICA ADICIONADA ---
    ' Salva o novo arquivo imediatamente com o nome e caminho corretos
    objWorkbook.SaveAs(caminhoArquivo)
    ' --- FIM DA CORREÇÃO ---
End If

' === FORMATAÇÃO DO CABEÇALHO ===
Dim headers, headerRange
headers = Array("antigo", "novo", "fornecedor", "tp de contrato", "data do contrato", _
                "material", "quantidade", "orgz compra", "grupo", "fim da validade", _
                "condicao de pgt", "referencia")
Set headerRange = objSheet.Range("A1:L1")
headerRange.Value = headers
headerRange.Font.Bold = True
headerRange.Interior.Color = RGB(144, 238, 144)
objSheet.Range("B1").Interior.Color = RGB(173, 216, 230)
objSheet.Columns("A:L").AutoFit

' === INPUT UNIFICADO COM VALIDAÇÃO ===
Dim entrada, partes, V_Data_INI, V_Data_FIM, V_TP_PEDIDO, V_Quantidade
entrada = InputBox("Digite os parâmetros separados por vírgula:" & vbCrLf & _
                   "Data Início, Data Fim, Tipo Pedido, Quantidade" & vbCrLf & _
                   "Exemplo: 01.01.2024,01.12.2024,ZD,10")

If entrada = "" Then
    MsgBox "Entrada cancelada pelo usuário."
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit
End If

partes = Split(entrada, ",")
If UBound(partes) <> 3 Then
    MsgBox "Por favor, preencha exatamente 4 valores separados por vírgula."
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit
End If

V_Data_INI = Trim(partes(0))
V_Data_FIM = Trim(partes(1))
V_TP_PEDIDO = Trim(partes(2))
V_Quantidade = Trim(partes(3))

' === ACESSA A TABELA EKKO VIA SE16N ===
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,8]").text = V_Data_INI
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,8]").text = V_Data_FIM
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = V_TP_PEDIDO
session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = V_Quantidade
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
If Not session.findById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
On Error GoTo 0

session.findById("wnd[0]/tbar[1]/btn[8]").press

' === CAPTURA DE DADOS DO GRID ===
Dim grid, startTime, maxWaitTime
maxWaitTime = 120 ' Timeout de 2 minutos.
startTime = Timer

Do
    On Error Resume Next
    Set grid = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    On Error GoTo 0

    If Not grid Is Nothing Then
        Exit Do
    End If

    If Timer - startTime > maxWaitTime Then
        MsgBox "A tabela de resultados não carregou em " & maxWaitTime & " segundos. O script será encerrado."
        objWorkbook.Close False
        objExcel.Quit
        WScript.Quit
    End If

    WScript.Sleep 500
Loop

Dim totalLinhas, linhaExcel, V_DocNum, i
linhaExcel = 2
totalLinhas = grid.RowCount

For i = 0 To totalLinhas - 1
    If i Mod 42 = 0 Then
        grid.FirstVisibleRow = i
        WScript.Sleep 300
    End If
    
    V_DocNum = grid.GetCellValue(i, "EBELN")
    objSheet.Cells(linhaExcel, 1).Value = V_DocNum
    
    linhaExcel = linhaExcel + 1
Next

' === FINALIZAÇÃO ===
If Not objWorkbook Is Nothing Then
    objWorkbook.Save
    objSheet.Columns("A:L").AutoFit
End If

MsgBox "Dados copiados para o Excel com sucesso! Total de linhas: " & totalLinhas

' Limpa os objetos da memória
Set grid = Nothing
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set fso = Nothing