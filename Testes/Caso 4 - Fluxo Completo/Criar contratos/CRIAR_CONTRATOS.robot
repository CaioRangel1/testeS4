*** Settings ***
Documentation    Automação para criação de contratos no SAP via transação ME31K
...              Lê dados do Excel e cria contratos automaticamente
Library          Collections
Library          String
Library          RoboSAPiens
Library          RPA.Excel.Files

*** Variables ***
${EXCEL_FILE}         contratos.xlsx
${WORKSHEET_NAME}     Sheet1
${LINHA_INICIAL}      2

*** Test Cases ***
Criar Contratos SAP
    [Documentation]    Cria contratos no SAP ME31K baseado em dados do Excel
    [Tags]             sap    me31k    contratos    excel
    
    # Conecta ao SAP
    Connect To SAP
    
    # Abre Excel e processa contratos
    Open Excel And Process Contracts
    
    # Salva e fecha Excel
    Save And Close Excel
    
    Log    Criação de contratos finalizada com sucesso!

*** Keywords ***
Connect To SAP
    [Documentation]    Conecta ao SAP e maximiza a janela
    Connect To Server
    Maximize Sap Window

Open Excel And Process Contracts
    [Documentation]    Abre Excel e processa cada linha para criar contratos
    Open Workbook    ${EXCEL_FILE}
    
    # Obtém dados do Excel
    ${worksheet_data}=    Read Worksheet    name=${WORKSHEET_NAME}    header=True
    ${total_rows}=    Get Length    ${worksheet_data}
    
    Log    Total de linhas para processar: ${total_rows}
    
    # Processa cada linha (contrato)
    ${row_number}=    Set Variable    ${LINHA_INICIAL}
    FOR    ${row_data}    IN    @{worksheet_data}
        ${column_a_value}=    Get From Dictionary    ${row_data}    ${row_data.keys()[0]}    default=${EMPTY}
        
        # Para se célula da coluna A estiver vazia
        Run Keyword If    '${column_a_value}' == '${EMPTY}'    Exit For Loop
        
        # Extrai dados da linha atual
        ${contract_data}=    Extract Contract Data From Row    ${row_data}
        
        # Cria contrato no SAP
        ${new_contract_number}=    Create Contract    ${contract_data}
        
        # Atualiza Excel com número do novo contrato
        Write Table Cell    row=${row_number}    column=16    value=${new_contract_number}    name=${WORKSHEET_NAME}
        
        ${row_number}=    Evaluate    ${row_number} + 1
        Sleep    1s
    END

Extract Contract Data From Row
    [Documentation]    Extrai dados necessários de uma linha do Excel
    [Arguments]    ${row_data}
    
    ${values_list}=    Get Dictionary Values    ${row_data}
    
    # Mapeia colunas para variáveis (baseado na ordem das colunas do script VBS)
    ${contract_data}=    Create Dictionary
    ...    preco=${values_list}[1]              # Coluna B (2)
    ...    fornecedor=${values_list}[3]         # Coluna D (4)
    ...    cond_pagto=${values_list}[4]         # Coluna E (5)
    ...    org_compra=${values_list}[6]         # Coluna G (7)
    ...    grp_comprador=${values_list}[7]      # Coluna H (8)
    ...    material=${values_list}[9]           # Coluna J (10)
    ...    quantidade=${values_list}[10]        # Coluna K (11)
    ...    tipo_contrato=${values_list}[12]     # Coluna M (13)
    
    # Remove espaços em branco
    FOR    ${key}    IN    @{contract_data.keys()}
        ${value}=    Get From Dictionary    ${contract_data}    ${key}
        ${cleaned_value}=    Strip String    ${value}
        Set To Dictionary    ${contract_data}    ${key}    ${cleaned_value}
    END
    
    RETURN    ${contract_data}

Create Contract
    [Documentation]    Cria um contrato no SAP ME31K
    [Arguments]    ${contract_data}
    
    # Acessa transação ME31K
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/tbar[0]/okcd    /nme31k
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    # Preenche dados do cabeçalho
    ${fornecedor}=    Get From Dictionary    ${contract_data}    fornecedor
    ${tipo_contrato}=    Get From Dictionary    ${contract_data}    tipo_contrato
    ${org_compra}=    Get From Dictionary    ${contract_data}    org_compra
    ${grp_comprador}=    Get From Dictionary    ${contract_data}    grp_comprador
    
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-LIFNR    ${fornecedor}
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtRM06E-EVART    ${tipo_contrato}
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-EKORG    ${org_compra}
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-EKGRP    ${grp_comprador}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    # Captura data inicial e calcula data final
    ${data_inicial}=    Get Element Attribute    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-KDATB    text
    ${data_final}=    Calculate Final Date    ${data_inicial}
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-KDATE    ${data_final}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    # Preenche condições de pagamento
    ${cond_pagto}=    Get From Dictionary    ${contract_data}    cond_pagto
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-ZTERM    ${cond_pagto}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    # Preenche dados do item
    Fill Item Data    ${contract_data}
    
    # Configura condições do item
    Configure Item Conditions    ${contract_data}
    
    # Salva o contrato
    ${new_contract_number}=    Save Contract
    
    # Retorna ao menu principal
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/tbar[0]/okcd    /n
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    RETURN    ${new_contract_number}

Calculate Final Date
    [Documentation]    Calcula data final adicionando 1 ano à data inicial
    [Arguments]    ${data_inicial}
    
    # Separa dia, mês e ano
    ${partes}=    Split String    ${data_inicial}    .
    ${dia}=    Convert To Integer    ${partes}[0]
    ${mes}=    Convert To Integer    ${partes}[1]
    ${ano}=    Evaluate    ${partes}[2] + 1
    
    # Formata data final
    ${dia_formatado}=    Format String    {:02d}    ${dia}
    ${mes_formatado}=    Format String    {:02d}    ${mes}
    ${data_final}=    Set Variable    ${dia_formatado}.${mes_formatado}.${ano}
    
    RETURN    ${data_final}

Fill Item Data
    [Documentation]    Preenche dados do item do contrato
    [Arguments]    ${contract_data}
    
    ${material}=    Get From Dictionary    ${contract_data}    material
    ${quantidade}=    Get From Dictionary    ${contract_data}    quantidade
    ${preco}=    Get From Dictionary    ${contract_data}    preco
    
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMM06ETC_0220/ctxtEKPO-EMATN[3,0]    ${material}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-KTMNG[5,0]    ${quantidade}
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMM06ETC_0220/txtEKPO-NETPR[7,0]    ${preco}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]

Configure Item Conditions
    [Documentation]    Configura condições do item
    [Arguments]    ${contract_data}
    
    # Seleciona linha da tabela
    Select Table Row    id:/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMM06ETC_0220    0
    
    # Acessa aba de condições
    Click Button    id:/app/con[0]/ses[0]/wnd[0]/tbar[1]/btn[6]
    
    # Preenche condições de pagamento novamente
    ${cond_pagto}=    Get From Dictionary    ${contract_data}    cond_pagto
    Fill Text In Text Field    id:/app/con[0]/ses[0]/wnd[0]/usr/ctxtEKKO-ZTERM    ${cond_pagto}
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]

Save Contract
    [Documentation]    Salva o contrato e captura o número gerado
    
    # Salva o contrato
    Click Button    id:/app/con[0]/ses[0]/wnd[0]/tbar[0]/btn[11]
    Send VKey    0    /app/con[0]/ses[0]/wnd[0]
    
    # Confirma salvamento - primeiro popup
    ${popup_exists}=    Run Keyword And Return Status    Element Should Be Visible    id:/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1
    Run Keyword If    ${popup_exists}    Click Button    id:/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1
    
    # Confirma salvamento - segundo popup se existir
    ${popup2_exists}=    Run Keyword And Return Status    Element Should Be Visible    id:/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1
    Run Keyword If    ${popup2_exists}    Click Button    id:/app/con[0]/ses[0]/wnd[1]/usr/btnSPOP-OPTION1
    
    # Captura número do contrato da barra de status
    ${status_text}=    Get Element Attribute    id:/app/con[0]/ses[0]/wnd[0]/sbar    text
    ${contract_number}=    Get Substring    ${status_text}    -10
    
    RETURN    ${contract_number}

Save And Close Excel
    [Documentation]    Salva e fecha o arquivo Excel
    Save Workbook
    Close Workbook
    Log    Contratos criados com sucesso! Verifique a coluna 16 do Excel.
