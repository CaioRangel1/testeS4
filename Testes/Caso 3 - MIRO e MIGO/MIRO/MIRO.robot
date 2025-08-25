*** Settings ***
Documentation    Caso de teste para transação MIRO - Invoice Receipt
...              Convertido de script VBS para Robot Framework usando RoboSAPiens
...              
...              Este teste automatiza a criação de um recibo de fatura (Invoice Receipt) 
...              na transação MIRO do SAP, reproduzindo exatamente o comportamento do 
...              script VBS original.
Library          RoboSAPiens
Library          String
Library          DateTime
Library          OperatingSystem
Library          RPA.Tables
Library          RPA.Excel.Files

*** Variables ***
${COMPANY_CODE}         2000
${PAYMENT_METHOD}       0001
${PAYMENT_BLOCK}        A
${DOCUMENT_TYPE}        RE
${FISCAL_TYPE}          W1

*** Test Cases ***
Create Invoice Receipt MIRO
    [Documentation]    Executa o processo completo de criação de Invoice Receipt na transação MIRO
    [Tags]             MIRO    Invoice    Receipt    SAP    Pedido    Fatura    robot:recursive-continue-on-failure
    
    Prepare SAP
    
    Abrir Planilha de Dados de Teste    Dados apresentação 22-08.xlsx
    
    ${testData} =    Read Worksheet As Table    header=True
    FOR    ${index}    ${row}    IN ENUMERATE    @{testData}
        TRY
            Execute Transaction    /nmiro
            Configure Initial Data    ${row['NV PEDIDO']}
            Configure Payment Information
            Configure Details
            Configure Basic Data
            Configure Fiscal Information
            # Salvar o documento
            Press Key Combination    Ctrl+S
            Save MIRO return to Excel    ${{${index}+2}}
        EXCEPT
            ${statusbar}   Read Statusbar
            Log To Console    Erro ao processar pedido: ${row['NV PEDIDO']} - Mensagem de erro: ${statusbar['message']}
            # Fail    Erro ao processar pedido: ${row['NV PEDIDO']} - Mensagem de erro: ${statusbar['message']}
        END        
    END

*** Keywords ***
Prepare SAP
    [Documentation]    Conecta ao SAP
    Connect to Running SAP
    Maximize Window

Abrir Planilha de Dados de Teste
    [Documentation]    Abre o arquivo Excel com os dados para a execução dos casos de teste.
    [Arguments]    ${nomePlanilha}
    ${caminho_planilha} =    Join Path    ${CURDIR}    ../../..    Dados    ${nomePlanilha}
    Open Workbook    ${caminho_planilha}

Configure Initial Data
    [Documentation]    Configura os dados iniciais na aba inicial da MIRO
    [Arguments]    ${numPedido}
    
    # Preencher código da empresa na janela popup
    ${window_title} =    Get Window Title
    IF    '${window_title}' == 'Entrar empresa'
        Fill Text Field    Empresa    ${COMPANY_CODE}
        Press Key Combination    Enter
    END

    Fill Text Field    Nº do documento de compras    ${numPedido}
    Press Key Combination    Enter

    # Preencher data do documento (Data de hoje)
    ${curDate} =    Get Current Date    result_format=%d.%m.%Y
    Fill Text Field    Data no documento    ${curDate}
    Press Key Combination    Enter
    
    ${referencia} =    Generate Random String    6    [NUMBERS]
    
    # Preencher número de referência da fatura (sempre um N° aleatorio)
    Fill Text Field    Referência    ${referencia}

Configure Payment Information
    [Documentation]    Configura as informações de pagamento na aba Payment
    
    Select Tab    Pagamento
    
    Fill Text Field    Tipo de banco do parceiro    ${PAYMENT_METHOD}
    Select Dropdown Menu Entry    Bloq.pgto.    ${PAYMENT_BLOCK}
    Press Key Combination    Enter

Configure Details
    [Documentation]    Configura os detalhes do pedido de compra na aba Detalhe

    Select Tab    Detalhe
    
    Select Dropdown Menu Entry    Tp.doc.    ${DOCUMENT_TYPE}

    Fill Text Field    Ctg.NF    ${FISCAL_TYPE}
    Press Key Combination    Enter

Configure Basic Data
    [Documentation]    Configura os dados básicos na aba Basic Data
    
    Select Tab    DdsBásicos
    
    ${saldoDocumento} =    Read Text Field    Saldo do documento
    ${saldo} =    Strip String    ${saldoDocumento}    mode=RIGHT    characters=-
    Fill Text Field    Montante em moeda do documento    ${saldo}
    Press Key Combination    Enter

Configure Fiscal Information
    [Documentation]    Configura as informações fiscais na aba FI
    
    Select Tab    Imp.ret.fonte

    ${rowCount} =    Get Row Count
    FOR    ${index}    IN RANGE    1    ${rowCount}
        ${status}    ${cellValue} =    Run Keyword And Ignore Error    Read Cell    ${index}    Código IRF
        IF    $status == 'PASS' and $cellValue != '${EMPTY}'
            Fill Cell    ${index}    Código IRF    content=${EMPTY}
        END
    END

    Press Key Combination    Enter

Save MIRO return to Excel
    [Documentation]    Salva o retorno da transação MIRO no Excel.
    [Arguments]    ${index}
    ${statusbar}   Read Statusbar
    ${msgStatusBar} =    Set Variable    ${statusbar['message']}
    ${docMaterial} =    Evaluate    re.search(r'\\d{10}', $msgStatusBar).group(0)    modules=re
    Set Cell Value    ${index}    19    ${docMaterial}
    Save Workbook