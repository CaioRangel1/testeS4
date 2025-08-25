*** Settings ***
Library    OperatingSystem
Library    RoboSAPiens
Library    RPA.Tables
Library    RPA.Excel.Files

*** Variables ***
${STORAGE_LOCATION}    tran
${DELIVERY_NOTE}     22222222234
${DENSIDADE}    0,8600
${TEMP}     20,0

*** Test Cases ***
Executar MIGO
    [Documentation]    Executa a transação MIGO para entrada de mercadorias no SAP
    [Tags]    sap    migo
    
    Prepare SAP
    
    Abrir Planilha de Dados de Teste    Dados apresentação 22-08.xlsx
    ${testData}=    Read Worksheet As Table    header=True
    TRY
        FOR    ${row}    IN    @{testData}
            Log To Console    \nProcessando pedido: ${row['NV PEDIDO']}
            Execute Transaction    /nmigo
            Fill Purchase Order Details    ${row['NV PEDIDO']} 
            Configure Item Details
            Set Delivery Note
            Save Transaction
            Save MIGO return to Excel    ${row}
        END
    EXCEPT
        ${statusbar}   Read Statusbar
        Log To Console    Erro ao processar pedido: ${row['NV PEDIDO']} - Mensagem de erro: ${statusbar['message']}
        # Fail    ${statusbar['message']}
    END
    
*** Keywords ***
Abrir Planilha de Dados de Teste
    [Documentation]    Abre o arquivo Excel com os dados para a execução dos casos de teste.
    [Arguments]    ${nomePlanilha}
    ${caminho_planilha} =    Join Path    ${CURDIR}    ../../..    Dados    ${nomePlanilha}
    Open Workbook    ${caminho_planilha}

Prepare SAP
    [Documentation]    Conecta ao SAP
    Connect to Running SAP
    Maximize Window

Fill Purchase Order Details
    [Documentation]    Preenche o número do documento de compra.
    [Arguments]    ${docPedido}
    Fill Text Field    Nº do documento de compras    ${docPedido}
    Press Key Combination    Enter

Configure Item Details
    [Documentation]    Configura a localização de armazenamento, temperaturas e densidade do material.
    Select Tab    Od
    Fill Text Field    Depósito    ${STORAGE_LOCATION}
    Press Key Combination    Enter

    Select Tab    Quantidades adiciona

    ${materialTemperature} =    Read Cell    Material temperature.    Valor
    ${testTemperature} =    Read Cell    Test temperature.    Valor
    ${testDensity} =    Read Cell    Test density.    Valor

    IF    '${materialTemperature}' == '0,0'
        Fill Cell    Material temperature.    Valor    ${TEMP}
    END

    IF    '${testTemperature}' == '0,0'
        Fill Cell    Test temperature.    Valor    ${TEMP}
    END

    IF    '${testDensity}' == '0,0000'
        Fill Cell    Test density.    Valor    ${DENSIDADE}
    END

    Tick Checkbox    Item é transferido para o documento

Set Delivery Note
    [Documentation]    Preenche o número da nota de remessa.
    Fill Text Field    Nº nota de remessa externa    ${DELIVERY_NOTE}
    Press Key Combination    Enter

Save Transaction
    [Documentation]    Salva a transação MIGO.
    Press Key Combination    Ctrl+S
    ${statusbar}   Read Statusbar
    Log    MIGO Executada com sucesso. Documento de material: ${statusbar}
Save MIGO return to Excel
    [Documentation]    Salva o retorno da transação MIGO no Excel.
    [Arguments]    ${row}
    ${statusbar}   Read Statusbar
    ${msgStatusBar} =    Set Variable    ${statusbar['message']}
    ${docMaterial} =    Evaluate    re.search(r'\d{10}', $msgStatusBar).group(0)    modules=re
    Log To Console    MIGO: ${docMaterial}
    Set Cell Value    ${row}    MIGO    ${docMaterial}
    Save Workbook
