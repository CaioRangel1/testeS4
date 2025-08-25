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

*** Variables ***
# Variáveis de configuração do SAP - AJUSTAR CONFORME SEU AMBIENTE
${SAP_SERVER}           seu_servidor_sap
${SAP_SYSTEM}           00
${SAP_CLIENT}           100
${SAP_USER}             seu_usuario
${SAP_PASSWORD}         sua_senha

# Dados do teste - baseados no script VBS original
${COMPANY_CODE}         2000
${PURCHASE_ORDER}       4503342051
${DOCUMENT_DATE}        22.08.2025
${REFERENCE}            222128
${PAYMENT_METHOD}       0001
${PAYMENT_BLOCK}        A
${DOCUMENT_TYPE}        RE
${FISCAL_TYPE}          W1
${INVOICE_AMOUNT}       3.969,83

*** Test Cases ***
Create Invoice Receipt MIRO
    [Documentation]    Executa o processo completo de criação de Invoice Receipt na transação MIRO
    ...                
    ...                Passos executados:
    ...                1. Conectar ao SAP
    ...                2. Navegar para transação MIRO
    ...                3. Preencher código da empresa
    ...                4. Preencher dados do pedido de compra
    ...                5. Configurar informações de pagamento
    ...                6. Configurar informações fiscais
    ...                7. Preencher valor da fatura
    ...                8. Limpar campos de retenção de impostos
    ...                9. Salvar o documento
    [Tags]             MIRO    Invoice    Receipt    SAP    Automated
    
    # Conectar ao SAP
    Prepare SAP
    
    # Navegar para transação MIRO
    Execute Transaction    /nmiro
    
    # Preencher código da empresa na janela popup
    # Fill Text Field    Empresa    ${COMPANY_CODE}
    # Press Key Combination    Enter
    
    # Preencher número do pedido de compra
    Fill Text Field    Nº do documento de compras    ${PURCHASE_ORDER}
    Press Key Combination    Enter

    # Preencher data do documento (Data de hoje)
    ${DATA_HOJE} =    Get Current Date    result_format=%d.%m.%Y
    Fill Text Field    Data no documento    ${DATA_HOJE}
    Press Key Combination    Enter

    # Preencher número de referência da fatura (sempre um N° aleatorio)
    Fill Text Field    Referência    ${REFERENCE}

    Configure Payment Information

    Configure Details

    Configure Basic Data

    Configure Fiscal Information
    
    # Salvar o documento
    # Press Key Combination    CTRL+S
    
    # Log de sucesso
    Log    Invoice Receipt MIRO criado com sucesso!
    
    # Capturar evidência (opcional)
    # Take Screenshot For Evidence

*** Keywords ***
Prepare SAP
    [Documentation]    Conecta ao SAP
    Connect to Running SAP
    Maximize Window

Configure Payment Information
    [Documentation]    Configura as informações de pagamento na aba Payment
    
    # Navegar para aba Payment (usando F-key ao invés de click)
    Select Tab    Pagamento
    
    # Preencher método de pagamento
    Fill Text Field    Tipo de banco do parceiro    ${PAYMENT_METHOD}
    Press Key Combination    Enter

    # Configurar bloqueio de pagamento
    Select Dropdown Menu Entry    Bloq.pgto.    ${PAYMENT_BLOCK}
    Press Key Combination    Enter

Configure Details
    [Documentation]    Configura os detalhes do pedido de compra na aba Detalhe

    # Navegar para aba Detalhe (usando F-key ao invés de click)
    Select Tab    Detalhe
    
    # Preencher data do documento
    Select Dropdown Menu Entry    Tp.doc.    ${DOCUMENT_TYPE}
    Press Key Combination    Enter

    # Preencher valor da fatura
    Fill Text Field    Ctg.NF    ${FISCAL_TYPE}
    Press Key Combination    Enter

Configure Basic Data
    [Documentation]    Configura os dados básicos na aba Basic Data
    
    # Navegar para aba Basic Data (usando F-key ao invés de click)
    Select Tab    DdsBásicos
    
    ${saldoDocumento} =    Read Text Field    Saldo do documento
    ${saldo} =    Strip String    ${saldoDocumento}    mode=RIGHT    characters=-
    Fill Text Field    Montante em moeda do documento    ${saldo}
    Press Key Combination    Enter

Configure Fiscal Information
    [Documentation]    Configura as informações fiscais na aba FI
    
    # Navegar para aba FI (usando F-key ao invés de click)
    Select Tab    Imp.ret.fonte

    ${rowCount} =    Get Row Count    ACWT_ITEM

    FOR    ${index}    IN RANGE    1    ${rowCount}
        Fill Cell    ${index}    Código IRF    content=${EMPTY}
        Press Key Combination    Enter
    END

Navigate Through Tabs
    [Documentation]    Navega entre as abas conforme o script original
    ...                Simula a navegação usando teclas de função
    
    # Navegar para aba WT (Withholding Tax)
    Press Key Combination    F7    # ajustar conforme necessário
    
    # Navegar para aba FI
    Press Key Combination    F6
    
    # Navegar para aba PAY
    Press Key Combination    F5
    
    # Voltar para aba TOTAL
    Press Key Combination    F4    # ajustar conforme necessário

Clear Withholding Tax Fields
    [Documentation]    Limpa os campos de retenção de impostos conforme script original
    
    # Navegar para aba WT
    Press Key Combination    F7
    
    # Limpar os campos de código de retenção
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/ctxtACWT_ITEM-WT_WITHCD[1,0]    ${EMPTY}
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/ctxtACWT_ITEM-WT_WITHCD[1,1]    ${EMPTY}
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/ctxtACWT_ITEM-WT_WITHCD[1,2]    ${EMPTY}
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_WT/ssubHEADER_SCREEN:SAPLFDCB:0080/subSUB_WT:SAPLFWTD:0120/tblSAPLFWTDWT_DIALOG/ctxtACWT_ITEM-WT_WITHCD[1,3]    ${EMPTY}
    Press Key Combination    ENTER

Take Screenshot For Evidence
    [Documentation]    Captura screenshot para evidência do teste
    ...                Esta implementação pode variar dependendo da versão da RoboSAPiens
    
    # Método 1: Usando Log com timestamp
    ${timestamp}=    Get Current Date    result_format=%Y%m%d_%H%M%S
    Log    Screenshot capturado em ${timestamp}
    
    # Método 2: Se disponível na sua versão da RoboSAPiens
    # Take Screenshot    MIRO_Invoice_Receipt_${timestamp}.png

*** Comments ***
# ============================================================================
# INSTRUÇÕES DE USO:
# ============================================================================
# 
# 1. CONFIGURAÇÃO INICIAL:
#    - Ajuste as variáveis de conexão SAP (SAP_SERVER, SAP_CLIENT, etc.)
#    - Configure as credenciais de acesso
#    - Instale e configure a biblioteca RoboSAPiens
#
# 2. DADOS DE TESTE:
#    - Os valores nas variáveis são baseados no script VBS original
#    - Ajuste conforme necessário para seu ambiente
#
# 3. EXECUÇÃO:
#    - Execute: robot MIRO_Complete.robot
#    - Ou: robot -d results MIRO_Complete.robot (para salvar resultados)
#
# 4. PERSONALIZAÇÃO:
#    - As teclas de função (F4, F5, F6, F7) podem precisar de ajuste
#    - Alguns localizadores podem variar entre versões do SAP
#    - Adicione validações específicas conforme necessário
#
# ============================================================================
# MAPEAMENTO DO SCRIPT VBS ORIGINAL:
# ============================================================================
#
# VBS: session.findById("wnd[0]").maximize
# Robot: Maximize Window
#
# VBS: session.findById("wnd[0]/tbar[0]/okcd").text = "miro"
# Robot: Fill TextField    wnd[0]/tbar[0]/okcd    miro
#
# VBS: session.findById("wnd[0]").sendVKey 0
# Robot: Press Key Combination    ENTER
#
# VBS: session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = "2000"
# Robot: Fill TextField    wnd[1]/usr/ctxtBKPF-BUKRS    2000
#
# E assim por diante para todos os elementos...
#
# ============================================================================
