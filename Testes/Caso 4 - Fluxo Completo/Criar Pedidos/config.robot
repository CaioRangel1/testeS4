*** Settings ***
Documentation    Configurações e variáveis para os testes de contratos SAP

*** Variables ***
# === CONFIGURAÇÕES DO EXCEL ===
${EXCEL_FILE}         contratos.xlsx
${WORKSHEET_NAME}     Sheet1
${LINHA_INICIAL}      2

# === MAPEAMENTO DE COLUNAS DO EXCEL ===
# Coluna A (1)  - Identificador/Controle
# Coluna B (2)  - Preço
# Coluna C (3)  - [Não usado]
# Coluna D (4)  - Fornecedor
# Coluna E (5)  - Condições de Pagamento
# Coluna F (6)  - [Não usado]
# Coluna G (7)  - Organização de Compras
# Coluna H (8)  - Grupo de Comprador
# Coluna I (9)  - [Não usado]
# Coluna J (10) - Material
# Coluna K (11) - Quantidade
# Coluna L (12) - [Não usado]
# Coluna M (13) - Tipo de Contrato
# Coluna P (16) - Número do Contrato Criado (saída)

# === CONFIGURAÇÕES SAP ===
${SAP_SERVER}         SAP_SERVER_NAME
${SAP_CLIENT}         100
${SAP_USER}           SAP_USER
${SAP_PASSWORD}       SAP_PASSWORD

# === TRANSAÇÕES SAP ===
${TRANSACAO_CONTRATO}     /nme21n

# === TIMEOUTS ===
${DEFAULT_TIMEOUT}    30s
${POPUP_TIMEOUT}      10s
