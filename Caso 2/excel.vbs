' === INICIA EXCEL ===
Dim objExcel, wb, ws
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set wb = objExcel.Workbooks.Add
Set ws = wb.Sheets(1)
ws.Name = "Tabela Contratos"

' === CABEÇALHOS COM ACENTOS ===
Dim headers, i
headers = Array("antigo", "novo", "n contrato", "fornecedor", "tp de contrato", "data do contrato", _
                "material", "quantidade", "orgz compra", "grupo", "fim da validade", "condição de pgt", "referência")

' === INSERE CABEÇALHOS E ESTILO ===
For i = 0 To UBound(headers)
    With ws.Cells(1, i + 1)
        .Value = headers(i)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 11

        ' === CORES ===
        If headers(i) = "novo" Then
            .Interior.Color = RGB(0, 112, 192)   ' Azul
            .Font.Color = RGB(255, 255, 255)     ' Letra branca
        Else
            .Interior.Color = RGB(198, 239, 206) ' Verde claro
        End If
    End With

    ' === LARGURA PERSONALIZADA ===
    If i = 0 Or i = 1 Then ' "antigo" ou "novo"
        ws.Columns(i + 1).ColumnWidth = 11.43
    Else
        ws.Columns(i + 1).AutoFit
    End If
Next

