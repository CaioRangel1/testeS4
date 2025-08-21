*** Settings ***
Library    OperatingSystem
Library    RoboSAPiens
Library    RPA.Tables
Library    RPA.Excel.Files

*** Variables ***
${STORAGE_LOCATION}  tran
${DELIVERY_NOTE}     22222222234

*** Test Cases ***
Executar MIGO
    [Documentation]    Executa a transação MIGO para entrada de mercadorias no SAP
    [Tags]    sap    migo

    Prepare SAP

    Abrir Planilha de Dados de Teste    Dados apresentação 22-08.xlsx
    ${testData}=    Read Worksheet As Table    header=True

    FOR    ${row}    IN    @{testData}
        Log To Console    Processando pedido: ${row['Pedido Origem']}
        Execute Transaction    /nmigo
        #Trocar para 'Pedido Novo'
        Fill Purchase Order Details    ${row['Pedido Origem']} 
        Configure Item Details
        Set Delivery Note
        Save Transaction
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
    [Documentation]    Preenche o número do documento de compra
    [Arguments]    ${docPedido}
    Fill Text Field    Nº do documento de compras    ${docPedido}
    Press Key Combination    Enter

Configure Item Details
    [Documentation]    Configura a localização de armazenamento e a caixa de seleção de detalhes
    Select Tab    Od
    Fill Text Field    Depósito    ${STORAGE_LOCATION}
    Press Key Combination    Enter

    Select Tab    Quantidades adiciona
    Tick Checkbox    Item é transferido para o documento

Set Delivery Note
    [Documentation]    Preenche o número da nota de remessa
    Fill Text Field    Nº nota de remessa externa    ${DELIVERY_NOTE}
    Press Key Combination    Enter

Save Transaction
    [Documentation]    Salva a transação MIGO
    # Push Button    Registrar
    Press Key Combination    Ctrl+S
    ${statusbar}   Read Statusbar
    Log    MIGO Executada com sucesso. Documento de material: ${statusbar}
