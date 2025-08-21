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

' === CRIA E CONFIGURA O EXCEL ===
Dim objExcel, objWorkbook, objSheet
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True 'Torna o Excel visível
Set objWorkbook = objExcel.Workbooks.Add
Set objSheet = objWorkbook.Sheets(1)

' === CRIAÇÃO E FORMATAÇÃO DO CABEÇALHO ===
' -- Define os títulos das colunas
Dim headers(15)
headers(0) = "Pedido Origen"
headers(1) = "Novo Pedido"
headers(2) = "TIP DE PEDIDO"
headers(3) = "FORNECEDOR"
headers(4) = "COND.PAG"
headers(5) = "INCOTERM"
headers(6) = "ORG. COMPRAS"
headers(7) = "GP. COMPRADOR"
headers(8) = "EMPRESA"
headers(9) = "MATERIAL"
headers(10) = "QNTD"
headers(11) = "CENTRO"
headers(12) = "CONTRATO"
headers(13) = "ITEM CONT."
headers(14) = "Data de Remessa"
headers(15) = "Status"

' -- Aplica os títulos e a formatação
Dim headerRange
Set headerRange = objSheet.Range("A1:P1")

' -- Escreve os cabeçalhos na planilha
For i = 0 To UBound(headers)
    objSheet.Cells(1, i + 1).Value = headers(i)
Next

' -- Formatação geral do cabeçalho
headerRange.Font.Bold = True
headerRange.Font.Color = vbWhite

' -- Cores de fundo
objSheet.Range("A1:O1").Interior.Color = RGB(0, 112, 192) ' Azul
objSheet.Range("P1").Interior.Color = RGB(0, 176, 80) ' Verde
objSheet.Range("B1").Interior.Color = RGB(0, 176, 80) 

' -- AutoAjuste das colunas
headerRange.Columns.AutoFit

' === INPUT UNIFICADO COM VALIDAÇÃO ===
Dim entrada, partes, V_Data_INI, V_Data_FIM, V_TP_PEDIDO
entrada = InputBox("Digite os parâmetros separados por vírgula:" & vbCrLf & _
                   "Data Início, Data Fim, Tipo Pedido" & vbCrLf & _
                   "Exemplo: 01.01.2024,01.12.2024,ZD")

If entrada = "" Then
    MsgBox "Entrada cancelada pelo usuário."
    objExcel.Quit ' Fecha o Excel se a operação for cancelada
    Set objExcel = Nothing
    WScript.Quit
End If

partes = Split(entrada, ",")

If UBound(partes) <> 2 Then
    MsgBox "Por favor, preencha exatamente 3 valores separados por vírgula."
    objExcel.Quit ' Fecha o Excel se a entrada for inválida
    Set objExcel = Nothing
    WScript.Quit
End If

V_Data_INI = Trim(partes(0))
V_Data_FIM = Trim(partes(1))
V_TP_PEDIDO = Trim(partes(2))

' === ACESSA A TABELA EKKO VIA SE16N ===
session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16n"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtGD-TAB").text = "EKKO"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,8]").text = V_Data_INI
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-HIGH[3,8]").text = V_Data_FIM
session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,4]").text = V_TP_PEDIDO

session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = "99999999"
session.findById("wnd[0]").sendVKey 0

On Error Resume Next
If Not session.findById("wnd[1]/tbar[0]/btn[0]") Is Nothing Then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
End If
On Error GoTo 0

session.findById("wnd[0]/tbar[1]/btn[8]").press

' === CAPTURA DE DADOS DO GRID ===
Dim grid, totalLinhas, linhaExcel, V_DocNum, i
Set grid = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")

linhaExcel = 2 ' Começa na linha 2 do Excel pois a linha 1 é o cabeçalho
totalLinhas = grid.RowCount

For i = 0 To totalLinhas - 1
    If i Mod 42 = 0 Then
        grid.FirstVisibleRow = i
        WScript.Sleep 300 ' Espera o SAP renderizar as linhas
    End If

    V_DocNum = grid.GetCellValue(i, "EBELN")
    objSheet.Cells(linhaExcel, 1).Value = V_DocNum ' Coluna A: "Pedido Origen"
    linhaExcel = linhaExcel + 1
Next

MsgBox "Dados copiados para o Excel com sucesso! Total de linhas: " & linhaExcel - 2

' === LIBERA OBJETOS ===
Set grid = Nothing
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing