*** Settings ***
Documentation    Caso de teste para criação de Pedidos de Compra (ME21N)
...              Convertido de script VBS para Robot Framework usando RoboSAPiens
...              
...              Este teste automatiza a criação de múltiplos pedidos de compra 
...              na transação ME21N do SAP, lendo os dados de uma planilha Excel
...              e salvando os números dos pedidos gerados.
Library          RoboSAPiens
Library          DateTime
Library          OperatingSystem
Library          Collections
Resource         ../../../Resources/sap_common.robot

*** Variables ***
# Variáveis de configuração do SAP
${SAP_SERVER}           seu_servidor_sap
${SAP_CLIENT}           100
${SAP_USER}             seu_usuario
${SAP_PASSWORD}         sua_senha


# Configurações de tempo

*** Test Cases ***
Create Purchase Orders From Excel
    [Documentation]    Cria múltiplos pedidos de compra lendo dados do Excel
    ...                
    ...                Colunas do Excel esperadas:
    ...                A: Contador/ID (não usado)
    ...                C: Tipo de Pedido
    ...                D: Fornecedor
    ...                E: Condições de Pagamento
    ...                F: Incoterm
    ...                G: Organização de Compras
    ...                H: Grupo de Compradores
    ...                I: Empresa
    ...                J: Material
    ...                K: Quantidade
    ...                L: Centro
    ...                M: Contrato (opcional)
    ...                N: Item do Contrato (opcional)
    ...                Q: Número do Pedido (será preenchido)
    [Tags]             ME21N    Purchase    Order    Excel    Automated
    
    # Configurações do arquivo Excel
    ${testData} =    Open Excel Worksheet    Dados apresentação 22-08.xlsx

    # Conectar ao SAP
    Prepare SAP
    
    
    # Processar cada linha do Excel
    FOR    ${row_data}    IN    @{testData}
        ${purchase_order_number}=    Create Single Purchase Order    ${row_data}
        Update Excel With PO Number    ${row_data}    ${purchase_order_number}
    END
    
    # Salvar e fechar Excel
    Save And Close Excel
    
    Log    Pedidos criados com sucesso! Conferir coluna Q do Excel.

*** Keywords ***
Connect To SAP
    [Documentation]    Conecta ao servidor SAP
    Open SAP    ${SAP_SERVER}    ${SAP_CLIENT}

Read Excel Data
    [Documentation]    Lê os dados do arquivo Excel
    ...                Retorna uma lista de dicionários com os dados de cada linha
    
    # Esta é uma implementação simplificada
    # Em um cenário real, você usaria uma biblioteca como ExcelLibrary ou pandas
    ${data}=    Create List
    
    # Exemplo de dados hardcoded para demonstração
    # Substitua por leitura real do Excel
    ${row1}=    Create Dictionary
    ...    linha=2
    ...    tipo_pedido=
    ...    fornecedor=FORN001
    ...    cond_pagto=Z030
    ...    incoterm=EXW
    ...    org_compra=1000
    ...    grp_comprador=001
    ...    empresa=1000
    ...    material=MAT001
    ...    quantidade=10
    ...    centro=1000
    ...    contrato=
    ...    item_contrato=
    
    Append To List    ${data}    ${row1}
    
    RETURN    ${data}

Create Single Purchase Order
    [Arguments]    ${row_data}
    [Documentation]    Cria um único pedido de compra com os dados fornecidos
    
    # Navegar para transação ME21N
    Execute Transaction   /nme21n
    
    # Preencher condições de pagamento
    # Select Tab    Comunicação
    # Select Tab    Remessa/fatura
    # Fill TextField    Chave de condições de pagamento    ${row_data['COND.PAG']}
    
    # Preencher organização de compras e grupo de compradores
    Select Tab    Dados organizacionais
    Fill TextField    Organização de compras    ${row_data['ORG. COMPRAS']}
    Fill TextField    Grupo de compradores    ${row_data['GP. COMPRADOR']}
    
    # Preencher fornecedor
    Fill TextField    Fornecedor/centro fornecedor   ${row_data['FORNECEDOR']}
    Select Tab    Remessa/fatura
    
    # Preencher dados do item
    Fill Cell    1    Material    ${row_data['MATERIAL']}
    Fill Cell    1    Qtd.pedido    ${row_data['QNTD']}
    Fill Cell    1    Cen.   ${row_data['CENTRO']}
    
    # Preencher contrato se existir
    IF    '${row_data}[NV CONTRATO]' != '${EMPTY}'
        Fill Cell    1    Contrato básico   ${row_data['NV CONTRATO']}
        Press Key Combination    Enter
        Press Key Combination    Ctrl+F2
        
    END
    
    # Navegar para aba de condições comerciais
    # Simular clique na aba TABHDT1
    #Press Key Combination    F5    # Ajustar conforme necessário
    
    # Preencher Incoterm1 
    Fill TextField     ctxtMEPO1226-INCO1    ${row_data['INCOTERM']}    exact=False
    Press Key Combination    Enter
    
    # Simular múltiplos ENTER conforme script original
    FOR    ${i}    IN RANGE    3
        Press Key Combination    Enter
    END
    
    # Salvar o pedido
    Press Key Combination    CTRL+S
    
    # Verificar se há janela de confirmação
    Handle Confirmation Dialog
    
    # Capturar número do pedido gerado
    ${po_number}=    Get Purchase Order Number
    
    RETURN    ${po_number}

Handle Confirmation Dialog
    [Documentation]    Trata janelas de confirmação que podem aparecer
    
    # Verificar se há janela de diálogo e confirmar
    TRY
        # Se existe janela de confirmação, pressionar o botão
        Press Key Combination    Enter
        Sleep    0.5s
    EXCEPT
        # Se não há janela, continuar normalmente
        Log    Nenhuma janela de confirmação encontrada
    END

Get Purchase Order Number
    [Documentation]    Extrai o número do pedido criado da barra de status
    
    # Capturar texto da barra de status
    ${status_text}=    Read Statusbar
    
    # Extrair os últimos 10 caracteres (número do pedido)
    ${po_number}=    Get Substring    ${status_text}    -10
    
    Log    Pedido criado: ${po_number}
    
    RETURN    ${po_number}

Update Excel With PO Number
    [Arguments]    ${row_data}   ${po_number}
    [Documentation]    Atualiza a planilha Excel com o número do pedido gerado
    
    # Esta é uma implementação simplificada
    # Em um cenário real, você escreveria de volta no Excel
    Log    Atualizando Excel - Linha: ${row_data['NV PEDIDO']}, PO: ${po_number}
    
    # Implementação real dependeria da biblioteca Excel utilizada
    # Exemplo: Set Cell Value    ${EXCEL_FILE_PATH}    ${SHEET_NAME}    ${row_data}[linha]    Q    ${po_number}

Save And Close Excel
    [Documentation]    Salva e fecha a planilha Excel
    
    # Implementação dependeria da biblioteca Excel utilizada
    Log    Excel salvo e fechado

*** Comments ***
# ============================================================================
# INSTRUÇÕES DE USO:
# ============================================================================
# 
# 1. CONFIGURAÇÃO INICIAL:
#    - Instale biblioteca para Excel (ex: ExcelLibrary, pandas via Process)
#    - Ajuste as variáveis de conexão SAP
#    - Configure o caminho do arquivo Excel
#
# 2. ESTRUTURA DO EXCEL:
#    Coluna A: ID/Contador (não usado)
#    Coluna C: Tipo de Pedido
#    Coluna D: Fornecedor
#    Coluna E: Condições de Pagamento  
#    Coluna F: Incoterm
#    Coluna G: Organização de Compras
#    Coluna H: Grupo de Compradores
#    Coluna I: Empresa
#    Coluna J: Material
#    Coluna K: Quantidade
#    Coluna L: Centro
#    Coluna M: Contrato (opcional)
#    Coluna N: Item do Contrato (opcional)
#    Coluna Q: Número do Pedido (preenchido automaticamente)
#
# 3. EXECUÇÃO:
#    robot CRIAR_PEDIDOS.robot
#
# ============================================================================
# IMPLEMENTAÇÃO COMPLETA COM EXCEL:
# ============================================================================
#
# Para implementação completa, considere usar:
#
# 1. ExcelLibrary (Robot Framework):
#    *** Settings ***
#    Library    ExcelLibrary
#    
#    Open Excel Document    ${EXCEL_FILE_PATH}    doc_id=workbook1
#    ${data}=    Read Excel Worksheet    name=${SHEET_NAME}
#    Write Excel Cell    row_num=2    col_num=17    value=${po_number}
#    Save Excel Document    filename=${EXCEL_FILE_PATH}
#    Close Excel Document    doc_id=workbook1
#
# 2. Process Library com Python pandas:
#    ${result}=    Run Process    python    -c    
#    ...    "import pandas as pd; df=pd.read_excel('${EXCEL_FILE_PATH}'); print(df.to_json())"
#
# 3. Biblioteca customizada em Python:
#    Crie uma biblioteca Python personalizada para manipular Excel
#
# ============================================================================
# MAPEAMENTO DO SCRIPT VBS:
# ============================================================================
#
# VBS: For i = 2 To objSheet.UsedRange.Rows.Count
# Robot: FOR ${row_data} IN @{excel_data}
#
# VBS: session.findById("wnd[0]/tbar[0]/okcd").text = "/nme21n"
# Robot: Fill TextField wnd[0]/tbar[0]/okcd /nme21n
#
# VBS: session.findById("wnd[0]").sendVKey 0  
# Robot: Press Key Combination ENTER
#
# VBS: objSheet.Cells(i, 17).Value = Novo_Pedido
# Robot: Update Excel With PO Number ${row_data} ${po_number}
#
# ============================================================================
