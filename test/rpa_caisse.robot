*** Settings ***
Documentation     Automatisation du processus d'insertion de donnée dans une base d'un fichier excel
Suite Setup       Connect To Database    psycopg2    ${DBName}    ${DBUser}    ${DBPass}    ${DBHost}    ${DBPort}
Suite Teardown    Disconnect From Database
Library           DatabaseLibrary
Library           ExcelLibrary
Library           BuiltIn


*** Variables ***
${DBHost}         localhost
${DBName}         rpa_caisse_db
${DBPass}         postgres
${DBPort}         5432
${DBUser}         postgres

${chemin}         ${CURDIR}${/}..\\ressources\\donneecaisse.xlsx
${feuille}        rapport


*** Keywords ***
Lire le fichier Excel
    [Documentation]        Lire le fichier excel et recuperer les données neccessaires
    [Arguments]                     ${chemin}       ${feuille}      ${ligne}    ${colonne}
    Open Excel Document             ${chemin}           1
    Get Sheet                       ${feuille}
    ${value}    Read Excel Cell     ${ligne}        ${colonne}
    [Return]                        ${value}
    Close All Excel Documents

Insertion dans la base de donnee
    [Documentation]      Insertion des données dans la base de donnée
    [Arguments]    ${nom_responsable}    ${email_responsable}    ${date}    ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${query}    Catenate       INSERT INTO  rapport_journalier (nom_responsable, email_responsable, date, montant_carte_bancaire, montant_espece, montant_ticket_restaurant, montant_prelevement, montant_monnaie ) VALUES ('${nom_responsable}','${email_responsable}','${date}','${montant_carte_bancaire}','${montant_espece}','${montant_ticket_restaurant}','${montant_prelevement}','${montant_apport_monnaie}')
    Execute Sql String    ${query}

Vérification des montants
    [Documentation]     Vérification du solde selon la regles "carte bancaire + espèces + ticket restaurant = prélèvement - apport monnaie".
    [Arguments]         ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${montant_total}    Evaluate            ${montant_carte_bancaire}+${montant_espece}+${montant_ticket_restaurant}
    ${solde}            Evaluate            ${montant_prelevement} - ${montant_apport_monnaie}
    ${solde_valid}      Run Keyword If      '${montant_total}'=='${solde}'    Set Variable    ${True}     ELSE    Set Variable    ${False}
    [Return]            ${solde_valid}



*** Test Cases ***
RPA CAISSE
    ${nom_responsable}              Lire le fichier Excel       ${chemin}       ${feuille}       3       3
    ${email_responsable}            Lire le fichier Excel       ${chemin}       ${feuille}       4       3
    ${date}                         Lire le fichier Excel       ${chemin}       ${feuille}       5       3
    ${montant_carte_bancaire}       Lire le fichier Excel       ${chemin}       ${feuille}       11      4
    ${montant_espece}               Lire le fichier Excel       ${chemin}       ${feuille}       12      4
    ${montant_ticket_restaurant}    Lire le fichier Excel       ${chemin}       ${feuille}       13      4
    ${montant_prelevement}          Lire le fichier Excel       ${chemin}       ${feuille}       15      4
    ${montant_apport_monnaie}       Lire le fichier Excel       ${chemin}       ${feuille}       16      4

    Insertion dans la base de donnee        ${nom_responsable}          ${email_responsable}    ${date}    ${montant_carte_bancaire}   ${montant_espece}    ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}
    ${solde}        Vérification des montants               ${montant_carte_bancaire}   ${montant_espece}       ${montant_ticket_restaurant}   ${montant_prelevement}   ${montant_apport_monnaie}

