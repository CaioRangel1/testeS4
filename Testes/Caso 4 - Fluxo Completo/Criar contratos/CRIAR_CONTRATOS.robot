*** Settings ***
Documentation    Automação para criação de contratos no SAP via transação ME31K
...              Lê dados do Excel e cria contratos automaticamente
Library          Collections
Library          String
Library          RoboSAPiens
Library          RPA.Excel.Files
Resource         ../../../Resources/sap_common.robot

*** Variables ***
${EXCEL_FILE}         contratos.xlsx
${WORKSHEET_NAME}     Sheet1
${LINHA_INICIAL}      2

*** Test Cases ***
Criar Contratos SAP
    [Documentation]    Cria contratos no SAP ME31K baseado em dados do Excel
    [Tags]             sap    me31k    contratos    excel
    
    Prepare SAP

    ${testData} =    Open Excel Worksheet    Dados apresentação 22-08.xlsx
    
    Process Contracts    ${testData}
    
    Log    Criação de contratos finalizada com sucesso!

*** Keywords ***
Process Contracts
    [Documentation]    Processa cada linha do excel para criar contratos
    [Arguments]    ${testData}
    
    ${total_rows}=    Get Length    ${testData}
    
    Log    Total de linhas para processar: ${total_rows}
    
    ${row_number}=    Set Variable    ${LINHA_INICIAL}
    FOR    ${index}    ${row_data}    IN ENUMERATE    @{testData}        
        # Cria contrato no SAP
        ${new_contract_number}=    Create Contract    ${row_data}
        
        Save Return to Excel    ${row_number}    ${new_contract_number}
        
        ${row_number}=    Evaluate    ${row_number} + 1
    END

Create Contract
    [Documentation]    Cria um contrato no SAP pela transação ME31K
    [Arguments]    ${contract_data}
    
    # Preenche dados iniciais
    Fill Text Field    Nº conta do fornecedor    ${contract_data['FORNECEDOR']}
    Fill Text Field    Tipo de contrato    ${contract_data['TP CONTRATO']}
    Fill Text Field    Organização de compras    ${contract_data['ORG. COMPRAS']}
    Fill Text Field    Grupo de compradores    ${contract_data['GP. COMPRADOR']}
    Press Key Combination    Enter
    
    # Captura data inicial, calcula e preenche a data final
    ${data_inicial}=    Read Text Field    Início do período de validade
    ${data_final}=    Calculate Final Date    ${data_inicial}
    Fill Text Field    Fim da validade    ${data_final}
    
    Fill Text Field    Condições pgto.    ${contract_data['COND.PAG']}
    Press Key Combination    Enter
    
    # Preenche dados do item
    Fill Cell    1    Material    ${contract_data['MATERIAL']}
    Press Key Combination    Enter
    Fill Cell    1    Qtd.prev.    ${contract_data['QNTD']}
    Fill Cell    1    Preço líq.    ${contract_data['Preço']}
    Press Key Combination    Enter

    # Configura condições do item
    # Configure Item Conditions    ${contract_data}
    
    # Salva o contrato
    ${new_contract_number}=    Save Contract
    
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

Configure Item Conditions
    [Documentation]    Configura condições do item
    [Arguments]    ${contract_data}
    
    # Seleciona linha da tabela
    Select Table Row    1
    
    # Acessa aba de condições
    Press Key Combination    F6
    
    # Preenche condições de pagamento novamente
    Fill Text Field    Condições pgto.    ${contract_data['COND.PAG']}
    Press Key Combination    Enter

Save Contract
    [Documentation]    Salva o contrato e captura o número gerado
    
    # Salva o contrato
    Press Key Combination    Ctrl+S
    Select Text Field    Item
    Press Key Combination    Enter
    Press Key Combination    Enter

    ${window_title}=    Get Window Title
    # Confirma salvamento - primeiro popup
    IF    $window_title == 'Gravar doc.'
        Push Button    Sim
    END
    ${window_title2}=    Get Window Title
    IF    $window_title2 == 'Gravar doc.'
        Push Button    Sim
    END
    
    ${statusBar}=    Read Statusbar
    # Captura número do contrato da barra de status
    ${contract_number}=    Get Substring    ${statusBar['message']}    -10
    
    RETURN    ${contract_number}
Save Return to Excel
    [Documentation]    Salva o retorno da transação no Excel.
    [Arguments]    ${index}    ${contractNumber}
    Set Cell Value    ${index}    16    ${contractNumber}
    Save Workbook
Save Contract to Excel
    [Documentation]    Salva e fecha o arquivo Excel
    Save Workbook
    Close Workbook
    Log    Contratos criados com sucesso! Verifique a coluna 16 do Excel.
