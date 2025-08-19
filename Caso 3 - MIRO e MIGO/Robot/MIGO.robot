*** Settings ***
Library    RoboSAPiens

*** Variables ***
${PO_NUMBER}         4503341943
${STORAGE_LOCATION}  tran
${DELIVERY_NOTE}     22222222234

*** Test Cases ***
Execute MIGO Entrada de Mercadoria
    [Documentation]    Executa a transação MIGO para entrada de mercadorias no SAP
    [Tags]    sap    migo    goods_receipt
    
    Connect To SAP
    Execute Transaction    /nmigo
    Fill Purchase Order Details
    Configure Item Details
    Set Delivery Note
    Save Transaction

*** Keywords ***
Connect To SAP
    [Documentation]    Conecta ao SAP
    Connect to Running SAP
    Maximize Window

Fill Purchase Order Details
    [Documentation]    Preenche o número do documento de compra
    Fill Text Field    Nº do documento de compras    ${PO_NUMBER}
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
