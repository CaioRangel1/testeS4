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

' === PEGA O PRIMEIRO EXCEL ABERTO ===
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")

If objExcel Is Nothing Then
    MsgBox "Nenhum Excel encontrado aberto."
    WScript.Quit
End If

If objExcel.Workbooks.Count = 0 Then
    MsgBox "Excel está aberto, mas sem nenhuma planilha carregada."
    WScript.Quit
End If

Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
On Error GoTo 0


Set objWorkbook = objExcel.Workbooks(1)
Set objSheet = objWorkbook.Sheets(1)
On Error GoTo 0

' === LOOP NAS LINHAS DO EXCEL ===
Dim linha, fornecedor, tipoContrato, material, quantidade
Dim orgCompra, grupoCompra, condPag
linha = 2

Do While objSheet.Cells(linha, 1).Value <> ""
    On Error Resume Next

    preco          = Trim(objSheet.Cells(linha, 2).Value) 
    fornecedor     = Trim(objSheet.Cells(linha, 4).Value) 
    tipoContrato   = Trim(objSheet.Cells(linha, 13).Value) 
    material       = Trim(objSheet.Cells(linha, 10).Value) 
    quantidade     = Trim(objSheet.Cells(linha, 11).Value) 
    orgCompra      = Trim(objSheet.Cells(linha, 7).Value) 
    grupoCompra    = Trim(objSheet.Cells(linha, 8).Value) 
    condPag        = Trim(objSheet.Cells(linha, 5).Value)  

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
    session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").text = condPag
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]").text = material
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[5,0]").text = quantidade
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]").text = preco
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tblSAPMM06ETC_0220").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    session.findById("wnd[0]/usr/ctxtEKKO-ZTERM").text = condPag
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
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

    objSheet.Cells(linha, 16).Value = contratoNovo

ProximaLinha:
    linha = linha + 1
    On Error GoTo 0
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey 0

Loop

MsgBox "Criacao de contratos finalizada com sucesso!"
