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

' === LOOP PARA CADA PEDIDO (COLUNA A) ===
Dim i, pedido, linhaExcel
linhaExcel = 2

Do While objSheet.Cells(linhaExcel, 1).Value <> ""

    On Error Resume Next

    pedido = Trim(objSheet.Cells(linhaExcel, 1).Value)

    ' === ACESSA ME33K ===
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme33k"
    session.findById("wnd[0]").sendVKey 0

    ' Preenche o número do pedido
      session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").press
      session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").text = pedido
      session.findById("wnd[0]").sendVKey 0

    ' === EXTRAI DADOS ===
    Dim fornecedor, tipoDeContrato, dataContrato, material, qtdPrevia
    fornecedor = session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text
    tipoDeContrato = session.findById("wnd[0]/usr/ctxtRM06E-EVART").Text
    dataContrato = session.findById("wnd[0]/usr/ctxtRM06E-VEDAT").Text
    material = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").Text
    qtdPrevia = session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[6,0]").Text

    ' === TRATAMENTO DOS DADOS ===
    ' TODO: Melhorar o tratamento de strings
    If InStr(tipoDeContrato, " ") > 0 Then
       tipoDeContrato = Trim(Left(tipoDeContrato, InStr(tipoDeContrato, " ") - 1))
    Else
       tipoDeContrato = Trim(tipoDeContrato)
    End If

    If InStr(fornecedor, " ") > 0 Then
       fornecedor = Trim(Left(fornecedor, InStr(fornecedor, " ") - 1))
    Else
       fornecedor = Trim(fornecedor)
    End If

'----
' === ACESSA DETALHES ===
   session.findById("wnd[0]/tbar[1]/btn[6]").press

   Dim orgCompra, grupoCompradores, fimValidade, condPagamento, suaReferencia
   orgCompra = session.findById("wnd[0]/usr/ctxtEKKO-EKORG").Text
   grupoCompradores = session.findById("wnd[0]/usr/ctxtEKKO-EKGRP").Text
   fimValidade = session.findById("wnd[0]/usr/ctxtEKKO-KDATE").Text
   condPagamento = session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").Text
   suaReferencia = session.findById("wnd[0]/usr/txtEKKO-IHREZ").Text

    ' === ESCREVE NO EXCEL ===
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

    On Error GoTo 0

    linhaExcel = linhaExcel + 1
Loop

MsgBox "Extração finalizada com sucesso!"

