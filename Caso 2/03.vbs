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

linha = 2

Do While objSheet.Cells(linha, 1).Value <> ""

    fornecedor = Trim(objSheet.Cells(linha, 3).Value)
    tipoContrato = Trim(objSheet.Cells(linha, 4).Value)
    dataContrato = Trim(objSheet.Cells(linha, 5).Value)
    material = Trim(objSheet.Cells(linha, 6).Value)
    quantidade = Trim(objSheet.Cells(linha, 7).Value)
    orgCompra = Trim(objSheet.Cells(linha, 8).Value)
    grupoCompra = Trim(objSheet.Cells(linha, 9).Value)
    fimValidade = Trim(objSheet.Cells(linha, 10).Value)
    condPag = Trim(objSheet.Cells(linha, 11).Value)
    referencia = Trim(objSheet.Cells(linha, 12).Value)

    ' === AJUSTA DATA DE VALIDADE (+1 ano) ===
    If fimValidade <> "" Then
        partesData = Split(fimValidade, ".")
        If UBound(partesData) = 2 Then
            dia = CInt(partesData(0))
            mes = CInt(partesData(1))
            ano = CInt(partesData(2)) + 1
            novaData = Right("0" & dia, 2) & "." & Right("0" & mes, 2) & "." & ano
        Else
            MsgBox "Linha " & linha & ": Formato de data inválido: " & fimValidade
            objSheet.Cells(linha, 2).Value = "ERRO - Formato inválido"
            linha = linha + 1
        End If
    Else
        MsgBox "Linha " & linha & ": Data de validade está vazia."
        objSheet.Cells(linha, 2).Value = "ERRO - Data vazia"
        linha = linha + 1
    End If
    ' === ACESSA ME31K ===
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme31k"
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text = fornecedor
    session.findById("wnd[0]/usr/ctxtRM06E-EVART").text = tipoContrato
    session.findById("wnd[0]/usr/ctxtEKKO-EKORG").text = orgCompra
    session.findById("wnd[0]/usr/ctxtEKKO-EKGRP").text = grupoCompra
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtEKKO-KDATE").text = novaData
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").text = condPag
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtEKKO-IHREZ").text = referencia
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").text = material
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[5,0]").text = quantidade

    session.findById("wnd[0]").sendVKey 11 ' salvar

    ' === CONFIRMAÇÃO E CAPTURA DO NÚMERO ===
    On Error Resume Next
    If session.Children.Count > 1 Then
        If session.Children(1).Type = "GuiModalWindow" Then
            session.Children(1).FindById("usr/btnSPOP-OPTION1").press
        End If
    End If
    On Error GoTo 0

    msg = session.findById("wnd[0]/sbar").Text
    pedidoNovo = ""

    For i = 1 To Len(msg)
        ch = Mid(msg, i, 1)
        If IsNumeric(ch) Then
            pedidoNovo = pedidoNovo & ch
        End If
    Next

    objSheet.Cells(linha, 2).Value = pedidoNovo

ProximaLinha:
    linha = linha + 1
Loop

MsgBox "Processo finalizado com sucesso!"
