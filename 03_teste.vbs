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

' === LOOP NAS LINHAS DO EXCEL ===
Dim linha, fornecedor, tipoContrato, material, quantidade
Dim orgCompra, grupoCompra, condPag, referencia
linha = 2

Do While objSheet.Cells(linha, 1).Value <> ""

    On Error Resume Next

    fornecedor     = Trim(objSheet.Cells(linha, 3).Value)
    tipoContrato   = Trim(objSheet.Cells(linha, 4).Value)
    material       = Trim(objSheet.Cells(linha, 6).Value)
    quantidade     = Trim(objSheet.Cells(linha, 7).Value)
    orgCompra      = Trim(objSheet.Cells(linha, 8).Value)
    grupoCompra    = Trim(objSheet.Cells(linha, 9).Value)
    condPag        = Trim(objSheet.Cells(linha, 11).Value)
    referencia     = Trim(objSheet.Cells(linha, 12).Value)

    ' === ACESSA ME31K ===
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme31k"
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/usr/ctxtEKKO-LIFNR").text     = fornecedor
    session.findById("wnd[0]/usr/ctxtRM06E-EVART").text    = tipoContrato
    session.findById("wnd[0]/usr/ctxtEKKO-EKORG").text     = orgCompra
    session.findById("wnd[0]/usr/ctxtEKKO-EKGRP").text     = grupoCompra
    session.findById("wnd[0]").sendVKey 0

    ' === CAPTURA A DATA DE INÍCIO E ADICIONA UM ANO ===
    Dim dataInicialVal, novaDataVal, partes, dia, mes, ano
    dataInicialVal = session.findById("wnd[0]/usr/ctxtEKKO-KDATB").Text

    partes = Split(dataInicialVal, ".")
    If UBound(partes) = 2 Then
        dia = CInt(partes(0))
        mes = CInt(partes(1))
        ano = CInt(partes(2)) + 1
        novaDataVal = Right("0" & dia, 2) & "." & Right("0" & mes, 2) & "." & ano
        session.findById("wnd[0]/usr/ctxtEKKO-KDATE").text = novaDataVal
        session.findById("wnd[0]").sendVKey 0
    Else
        MsgBox "Linha " & linha & ": Erro ao ler data de início da validade (KDATB): " & dataInicialVal
        linha = linha + 1
    End If

    session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").text     = condPag
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtEKKO-IHREZ").text      = referencia
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").text = material
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[5,0]").text = quantidade
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 11  ' Salvar
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

    If session.Children.Count > 1 Then
        If session.Children(1).Type = "GuiModalWindow" Then
            session.Children(1).findById("usr/btnSPOP-OPTION1").press
        End If
    End If

    ' === CAPTURA O NÚMERO DO NOVO CONTRATO NA BARRA DE STATUS ===
    Dim contratoNovo
    contratoNovo = Right(Session.findById("wnd[0]/sbar").Text, 10)

    ' Dim msg, ch, contratoNovo, k
    ' msg = session.findById("wnd[0]/sbar").Text
    ' contratoNovo = ""
    ' MsgBox "Contrato criado: " & msg

    ' For k = 1 To Len(msg)
    '     ch = Mid(msg, k, 1)
    '     If IsNumeric(ch) Then
    '         contratoNovo = contratoNovo & ch
    '     End If
    ' Next

    objSheet.Cells(linha, 2).Value = contratoNovo

ProximaLinha:
    linha = linha + 1
    On Error GoTo 0

Loop

MsgBox "Criação de contratos finalizada com sucesso!"
