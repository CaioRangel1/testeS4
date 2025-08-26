*** Settings ***
Library          RoboSAPiens
Library          OperatingSystem
Library          String
Library          RPA.Tables
Library          RPA.Excel.Files

*** Variables ***
${USUARIO_PADRAO}    AUTOMATION_USER

*** Keywords ***
Prepare SAP
    [Documentation]    Conecta e prepara a instância do SAP
    Connect to Running SAP
    Maximize Window

Open Excel Worksheet
    [Documentation]    Abre o arquivo Excel com os dados para a execução dos casos de teste.
    [Arguments]    ${nomePlanilha}
    ${caminho_planilha} =    Join Path    ${CURDIR}    ../../..    Dados    ${nomePlanilha}
    Open Workbook    ${caminho_planilha}
    ${testData} =    Read Worksheet As Table    header=True

    RETURN    ${testData}