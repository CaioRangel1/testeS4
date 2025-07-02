' === CONEXÃO COM SAP ===
On Error Resume Next

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

If objSheet Is Nothing Then
    MsgBox "Planilha ativa não encontrada. Verifique se há uma pasta de trabalho aberta."
    WScript.Quit
End If

' === LOOP PARA CADA PEDIDO (COLUNA A) ===
For i = 2 To objSheet.UsedRange.Rows.Count   

    If Trim(objSheet.Cells(i, 1).Value) <> "" Then

        tipoPedido   = Trim(CStr(objSheet.Cells(i, 3).Value))
        fornecedor   = Trim(CStr(objSheet.Cells(i, 4).Value))
        condPagto    = Trim(CStr(objSheet.Cells(i, 5).Value))
        incoterm     = Trim(CStr(objSheet.Cells(i, 6).Value))
        orgCompra    = Trim(CStr(objSheet.Cells(i, 7).Value))
        grpComprador = Trim(CStr(objSheet.Cells(i, 8).Value))
        empresa      = Trim(CStr(objSheet.Cells(i, 9).Value))
        material     = Trim(CStr(objSheet.Cells(i, 10).Value))
        qtd          = Trim(CStr(objSheet.Cells(i, 11).Value))
        centro       = Trim(CStr(objSheet.Cells(i, 12).Value))
        contrato     = Trim(CStr(objSheet.Cells(i, 13).Value))
        itemContrato = Trim(CStr(objSheet.Cells(i, 14).Value))


        session.findById("wnd[0]/tbar[0]/okcd").text = "/nme21n"
        session.findById("wnd[0]").sendVKey 0

        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = fornecedor

        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").setFocus
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").key = tipoPedido

        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").text = orgCompra  
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").text = grpComprador
        session.findById("wnd[0]").sendVKey 0

        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-ZTERM").text = condPagto
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-INCO1").text = incoterm
        session.findById("wnd[0]").sendVKey 0

        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").text = material
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").text = qtd
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").text = centro
        session.findById("wnd[0]").sendVKey 0

    if contrato <> "" Then
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KONNR[27,0]").text = contrato
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-KTPNR[28,0]").text = itemContrato
        session.findById("wnd[0]").sendVKey 0
    End If

    For j = 1 To 6
        session.findById("wnd[0]").sendVKey 0
    Next

    session.findById("wnd[0]/tbar[0]/btn[11]").press

    If session.Children.Count > 1 Then
        session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
    End If

    Novo_Pedido = Right(Session.findById("wnd[0]/sbar").Text, 10)

    objSheet.Cells(i, 2).Value = Novo_Pedido

End If
WScript.Sleep 1000
Next

MsgBox "Pedidos criados com sucesso! Conferir coluna B do Excel."
