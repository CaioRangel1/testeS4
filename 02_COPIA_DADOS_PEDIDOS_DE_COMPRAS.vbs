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

    ' === ACESSA ME23N ===
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nme23n"
    session.findById("wnd[0]").sendVKey 0

    ' Preenche o número do pedido
      session.findById("wnd[0]/tbar[1]/btn[17]").press
      session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = pedido
      session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' === EXTRAI DADOS ===

    ' Tipo de pedido (BSART)
    Dim bsart
    bsart = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").text 

    ' Fornecedor (SUPERFIELD)
    Dim fornecedor
    fornecedor = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text

    ' === TRATAMENTO DOS DADOS ===

    If InStr(bsart, " ") > 0 Then
       bsart = Trim(Left(bsart, InStr(bsart, " ") - 1))
    Else
       bsart = Trim(bsart)
    End If

    If InStr(fornecedor, " ") > 0 Then
       fornecedor = Trim(Left(fornecedor, InStr(fornecedor, " ") - 1))
    Else
       fornecedor = Trim(fornecedor)
    End If

    ' Aba: Condição de pagamento
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1").select
    Dim zterm, inco1
    zterm = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-ZTERM").Text
    inco1 = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-INCO1").Text

    ' Aba: Organização de Compras
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9").select
    Dim ekorg, ekgrp, bukrs
    ekorg = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").Text
    ekgrp = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").Text
    bukrs = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").Text

    ' Itens da posição 0
    Dim ematn, menge, name1, konnr, ktpnr

    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").getAbsoluteRow(0).selected = true

'----
' === ACESSA A TABELA DE ITENS ===
Set grid = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/" & _
                            "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211")

    ematn = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").Text
    menge = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").Text
    name1 = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").Text
    konnr = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KONNR[27,0]").Text
    ktpnr = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-KTPNR[28,0]").Text
           
    ' Data de entrega
    Dim dataAtual
    dataAtual = Date
    Dim dataRemessa
    dataRemessa = DateAdd("d", 10, dataAtual)
    
'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").getAbsoluteRow(0).selected = true
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press

    'session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT5").select
    'eeind = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EEIND[9,0]").Text

    ' === ESCREVE NO EXCEL ===
      objSheet.Cells(linhaExcel, 3).Value = bsart
      objSheet.Cells(linhaExcel, 4).Value = fornecedor
      objSheet.Cells(linhaExcel, 5).Value = zterm
      objSheet.Cells(linhaExcel, 6).Value = inco1
      objSheet.Cells(linhaExcel, 7).Value = ekorg
      objSheet.Cells(linhaExcel, 8).Value = ekgrp
      objSheet.Cells(linhaExcel, 9).Value = bukrs
      objSheet.Cells(linhaExcel, 10).Value = ematn
      objSheet.Cells(linhaExcel, 11).Value = menge
      objSheet.Cells(linhaExcel, 12).Value = name1
      objSheet.Cells(linhaExcel, 13).Value = konnr
      objSheet.Cells(linhaExcel, 14).Value = ktpnr
      objSheet.Cells(linhaExcel, 15).Value = dataRemessa
      objSheet.Cells(linhaExcel, 16).Value = "OK"


    On Error GoTo 0

    linhaExcel = linhaExcel + 1
Loop

MsgBox "Extração finalizada com sucesso!"

