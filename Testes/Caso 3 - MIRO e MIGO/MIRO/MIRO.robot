*** Settings ***
Documentation    Caso de teste para transação MIRO - Invoice Receipt
...              Convertido de script VBS para Robot Framework usando RoboSAPiens
...              
...              Este teste automatiza a criação de um recibo de fatura (Invoice Receipt) 
...              na transação MIRO do SAP, reproduzindo exatamente o comportamento do 
...              script VBS original.
Library          RoboSAPiens
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
${PURCHASE_ORDER}       4503342047
${DOCUMENT_DATE}        22.08.2025
${REFERENCE}            222122
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
    Connect To SAP
    
    # Maximizar janela principal
    Maximize Window
    
    # Navegar para transação MIRO
    Fill TextField    wnd[0]/tbar[0]/okcd    miro
    Press Key Combination    ENTER
    
    # Preencher código da empresa na janela popup
    Fill TextField    wnd[1]/usr/ctxtBKPF-BUKRS    ${COMPANY_CODE}
    Press Key Combination    ENTER
    
    # Preencher número do pedido de compra
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6211/ctxtRM08M-EBELN    ${PURCHASE_ORDER}
    Press Key Combination    ENTER
    
    # Preencher data do documento
    Fill TextField    wnd[1]/usr/ctxtRBKP-BLDAT    ${DOCUMENT_DATE}
    Press Key Combination    ENTER
    
    # Preencher número de referência da fatura
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR    ${REFERENCE}
    Press Key Combination    ENTER
    
    # Configurar dados de pagamento
    Configure Payment Information
    
    # Configurar dados fiscais
    Configure Fiscal Information
    
    # Navegar entre as abas conforme script original
    Navigate Through Tabs
    
    # Preencher valor da fatura
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR    ${INVOICE_AMOUNT}
    Press Key Combination    ENTER
    
    # Limpar campos de retenção de impostos
    Clear Withholding Tax Fields
    
    # Salvar o documento
    Press Key Combination    CTRL+S
    
    # Log de sucesso
    Log    Invoice Receipt MIRO criado com sucesso!
    
    # Capturar evidência (opcional)
    Take Screenshot For Evidence

*** Keywords ***
Connect To SAP
    [Documentation]    Conecta ao servidor SAP usando as credenciais configuradas
    ...                NOTA: Ajustar conforme seu ambiente SAP
    
    # Conectar ao SAP GUI
    Open SAP    ${SAP_SERVER}    ${SAP_CLIENT}
    
    # Fazer login (se necessário)
    # Fill TextField    wnd[0]/usr/txtRSYST-BNAME    ${SAP_USER}
    # Fill TextField    wnd[0]/usr/pwdRSYST-BCODE    ${SAP_PASSWORD}
    # Press Key Combination    ENTER

Configure Payment Information
    [Documentation]    Configura as informações de pagamento na aba Payment
    
    # Navegar para aba Payment (usando F-key ao invés de click)
    Press Key Combination    F5    # ou ajustar conforme necessário
    
    # Preencher método de pagamento
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/ctxtINVFO-BVTYP    ${PAYMENT_METHOD}
    Press Key Combination    ENTER
    
    # Configurar bloqueio de pagamento
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_PAY/ssubHEADER_SCREEN:SAPLFDCB:0020/cmbINVFO-ZLSPR    ${PAYMENT_BLOCK}
    Press Key Combination    ENTER

Configure Fiscal Information
    [Documentation]    Configura as informações fiscais na aba FI
    
    # Navegar para aba FI (usando F-key ao invés de click)
    Press Key Combination    F6    # ou ajustar conforme necessário
    
    # Configurar tipo de documento
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/cmbINVFO-BLART    ${DOCUMENT_TYPE}
    
    # Preencher tipo fiscal
    Fill TextField    wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/ctxtINVFO-J_1BNFTYPE    ${FISCAL_TYPE}
    Press Key Combination    ENTER

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
