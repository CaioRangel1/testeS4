' === CONEXÃO COM SAP ===
If Not IsObject(application) Then
    Set SapGuiAuto  = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
    Set connection = application.Children(0)
End If
If Not IsObject(session) Then
    Set session = connection.Children(0)
End If
If IsObject(WScript) Then
    WScript.ConnectObject session,     "on"
    WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize

' === PEGA O PRIMEIRO EXCEL ABERTO ===
Dim objExcel, objWorkbook, objSheet
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If objExcel Is Nothing Then
    MsgBox "❌ O Excel não está aberto. Abre a planilha primeiro!"
    WScript.Quit
End If

If objExcel.Workbooks.Count = 0 Then
    MsgBox "❌ Nenhuma planilha aberta no Excel."
    WScript.Quit
End If

Set objWorkbook = objExcel.Workbooks(1)
Set objSheet = objWorkbook.Sheets(1)
On Error GoTo 0

' === COMEÇA O LOOP PRA CADA CONTRATO ===
Dim linhaExcel, pedido
linhaExcel = 2

Do While Trim(objSheet.Cells(linhaExcel, 1).Value) <> ""

    pedido = Trim(objSheet.Cells(linhaExcel, 1).Value)

    ' Vai pra ME33K
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme33k"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").text = pedido
    session.findById("wnd[0]").sendVKey 0

    ' Extrai dados principais
    Dim fornecedor, tipoDeContrato, dataContrato, material, qtdPrevia
    fornecedor = session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text
    tipoDeContrato = session.findById("wnd[0]/usr/ctxtRM06E-EVART").Text
    dataContrato = session.findById("wnd[0]/usr/ctxtRM06E-VEDAT").Text
    material = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").Text
    qtdPrevia = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[6,0]").Text

    ' Ajuste dos campos (tira os textos após espaços)
    If InStr(tipoDeContrato, " ") > 0 Then
        tipoDeContrato = Trim(Left(tipoDeContrato, InStr(tipoDeContrato, " ") - 1))
    End If
    If InStr(fornecedor, " ") > 0 Then
        fornecedor = Trim(Left(fornecedor, InStr(fornecedor, " ") - 1))
    End If

    ' Vai pra detalhes
    session.findById("wnd[0]/tbar[1]/btn[6]").press

    Dim orgCompra, grupoCompradores, fimValidade, condPagamento, suaReferencia
    orgCompra = session.findById("wnd[0]/usr/ctxtEKKO-EKORG").Text
    grupoCompradores = session.findById("wnd[0]/usr/ctxtEKKO-EKGRP").Text
    fimValidade = session.findById("wnd[0]/usr/ctxtEKKO-KDATE").Text
    condPagamento = session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").Text
    suaReferencia = session.findById("wnd[0]/usr/txtEKKO-IHREZ").Text

    ' Escreve no Excel
    objSheet.Cells(linhaExcel, 3).Value = fornecedor
    objSheet.Cells(linhaExcel, 4).Value = tipoDeContrato
    objSheet.Cells(linhaExcel, 5).Value = dataContrato
    objSheet.Cells(linhaExcel, 6).Value = material
    objSheet.Cells(linhaExcel, 7).Value = qtdPrevia
    objSheet.Cells(linhaExcel, 8).Value = orgCompra
    objSheet.Cells(linhaExcel, 9).Value = grupoCompradores
    objSheet.Cells(linhaExcel, 10).Value = fimValidade
    objSheet.Cells(linhaExcel, 11).Value = condPagamento
    objSheet.Cells(linhaExcel, 12).Value = suaReferencia

    linhaExcel = linhaExcel + 1
Loop

MsgBox "✅ Extração finalizada com sucesso!"
