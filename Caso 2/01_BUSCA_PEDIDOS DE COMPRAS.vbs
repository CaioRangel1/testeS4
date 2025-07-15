' === CONEXÃO COM SAP ===
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

' === CONEXÃO COM EXCEL ===
Dim objExcel, objSheet
Set objExcel = GetObject(, "Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

' === INPUT UNIFICADO COM VALIDAÇÃO ===
Dim entrada, partes, V_Data_INI, V_Data_FIM, V_TP_PEDIDO, V_Quantidade
entrada = InputBox("Digite os parâmetros separados por vírgula:" & vbCrLf & _
                   "Data Início, Data Fim, Tipo Pedido, Quantidade" & vbCrLf & _
                   "Exemplo: 01.01.2024,01.12.2024,ZD,10")

If entrada = "" Then
   MsgBox "Entrada cancelada pelo usuário."
   WScript.Quit
End If

partes = Split(entrada, ",")

If UBound(partes) <> 3 Then
   MsgBox "Por favor, preencha exatamente 4 valores separados por vírgula."
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
Dim grid, totalLinhas, linhaExcel, V_DocNum, i
Set grid = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")

linhaExcel = 2 ' Começa na linha 2 do Excel
'------------------------------------------------
totalLinhas = grid.RowCount

For i = 0 To totalLinhas - 1
    If i Mod 42 = 0 Then
        grid.FirstVisibleRow = i
        WScript.Sleep 300 ' Espera o SAP renderizar as linhas
    End If
   
    V_DocNum = grid.GetCellValue(i, "EBELN")
    objSheet.Cells(linhaExcel, 1).Value = V_DocNum
    linhaExcel = linhaExcel + 1
Next
'------------------------------------------------------------------

'For i = 0 To grid.RowCount - 1
'    V_DocNum = grid.GetCellValue(i, "EBELN")
'    objSheet.Cells(linhaExcel, 1).Value = V_DocNum
'    linhaExcel = linhaExcel + 1
'Next

MsgBox "Dados copiados para o Excel com sucesso! Total de linhas: " & linhaExcel - 2

