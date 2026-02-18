// =============================================================
// Google Apps Script pour le Questionnaire RSE - LACME SAS (FR)
// =============================================================
// Copiez tout ce code dans l'éditeur Google Apps Script.
//
// INSTRUCTIONS :
// 1. Ouvrez votre Google Sheet FR
// 2. Extensions > Apps Script
// 3. Collez ce code et sauvegardez
// 4. Exécutez la fonction "createHeaders" UNE SEULE FOIS
// 5. Déployez : Déployer > Nouveau déploiement > Application Web
//    - Exécuter en tant que : "Moi"
//    - Accès : "Tout le monde"
// 6. Copiez l'URL et collez-la dans index.html (SCRIPT_URL)
// =============================================================

function doPost(e) {
    try {
        var sheet = SpreadsheetApp.openById('1fxqGesXzS77-g5GV77p_tb_crSyfTf21ow4Hlidr0Qo').getActiveSheet();
        var data = e.parameter;
        var allData = e.parameters;

        // Helper pour les checkboxes (valeurs multiples)
        function getMulti(fieldName) {
            if (allData[fieldName]) {
                return allData[fieldName].join(', ');
            }
            return '';
        }

        // Helper pour les champs simples
        function getSingle(fieldName) {
            return data[fieldName] || '';
        }

        var row = [
            new Date(),                           // Horodatage
            getSingle('email'),                   // Q1
            getSingle('company_name'),            // Q2
            getSingle('address'),                 // Q3
            getSingle('siret'),                   // Q4
            getSingle('respondent_name'),         // Q5
            getSingle('respondent_title'),        // Q6
            getSingle('csr_contact'),             // Q7
            getSingle('structured_csr'),          // Q8
            // --- Branche A (Q9-Q39) ---
            getSingle('csr_labeled'),             // Q9
            getSingle('csr_label_details'),       // Q10
            getSingle('csr_signatory'),           // Q11
            getSingle('csr_signatory_details'),   // Q12
            getSingle('csr_responsible_exists'),  // Q13
            getSingle('csr_resp_name'),           // Q14
            getSingle('csr_resp_title'),          // Q15
            getSingle('csr_resp_email'),          // Q16
            getSingle('csr_report'),              // Q17
            getSingle('csr_report_link'),         // Q18
            getSingle('code_of_conduct'),         // Q19
            getSingle('whistleblowing'),          // Q20
            getSingle('human_rights_policy'),     // Q21
            getMulti('hr_areas'),                 // Q22
            getSingle('hr_other'),                // Q23
            getSingle('ohs_policy'),              // Q24
            getSingle('ohs_actions'),             // Q25
            getSingle('ohs_examples'),            // Q26
            getSingle('ethics_policy'),           // Q27
            getMulti('ethics_areas'),             // Q28
            getSingle('ethics_other'),            // Q29
            getSingle('env_policy'),              // Q30
            getSingle('env_system'),              // Q31
            getSingle('env_kpi'),                 // Q32
            getSingle('env_cert_details'),        // Q33
            getSingle('substances'),              // Q34
            getSingle('substances_proc'),         // Q35
            getSingle('supplier_csr'),            // Q36
            getMulti('supplier_comm'),            // Q37
            getSingle('supplier_other'),          // Q38
            getSingle('training_sessions'),       // Q39
            // --- Branche B (Q40-Q46) ---
            getSingle('informal_person'),         // Q40
            getSingle('informal_contact_details'),// Q41
            getMulti('basic_kpi'),                // Q42
            getSingle('basic_kpi_other'),         // Q43
            getMulti('written_rules'),            // Q44
            getSingle('written_rules_other'),     // Q45
            getSingle('support_interest'),        // Q46
            // --- Partie commune (Q47-Q63) ---
            getSingle('waste_measure'),           // Q47
            getSingle('waste_reduce'),            // Q48
            getSingle('waste_examples'),          // Q49
            getSingle('recycling'),               // Q50
            getSingle('recycling_types'),         // Q51
            getSingle('energy_measure'),          // Q52
            getSingle('energy_reduce'),           // Q53
            getSingle('energy_examples'),         // Q54
            getSingle('water_measure'),           // Q55
            getSingle('water_reduce'),            // Q56
            getSingle('water_examples'),          // Q57
            getSingle('transport_actions'),       // Q58
            getSingle('transport_examples'),      // Q59
            getSingle('co2_measure'),             // Q60
            getSingle('ecodesign'),               // Q61
            getSingle('ecodesign_products'),      // Q62
            getSingle('comments'),                // Q63
        ];

        sheet.appendRow(row);

        return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: error.toString() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

function doGet(e) {
    return ContentService.createTextOutput("Le script fonctionne correctement !")
        .setMimeType(ContentService.MimeType.TEXT);
}

// =============================================================
// Exécutez cette fonction UNE SEULE FOIS pour créer les en-têtes
// =============================================================
function createHeaders() {
    var sheet = SpreadsheetApp.openById('1fxqGesXzS77-g5GV77p_tb_crSyfTf21ow4Hlidr0Qo').getActiveSheet();

    var headers = [
        "Horodatage",
        "Email",
        "Nom de la Société",
        "Adresse",
        "SIRET",
        "Nom et Prénom du répondant",
        "Fonction du répondant",
        "Contact RSE/Environnement/Qualité",
        "Q8 - Démarche RSE structurée",
        "Q9 - RSE labellisée",
        "Q10 - Détails label",
        "Q11 - Signataire engagement volontaire",
        "Q12 - Détails engagement",
        "Q13 - Personne RSE dédiée",
        "Q14 - Nom personne RSE",
        "Q15 - Fonction personne RSE",
        "Q16 - Email personne RSE",
        "Q17 - Rapport RSE publié",
        "Q18 - Lien rapport RSE",
        "Q19 - Code de conduite",
        "Q20 - Système d'alerte",
        "Q21 - Politique Droits Humains",
        "Q22 - Domaines Droits Humains",
        "Q23 - Autres Droits Humains",
        "Q24 - Politique SST",
        "Q25 - Actions prévention risques",
        "Q26 - Exemples SST",
        "Q27 - Politique éthique des affaires",
        "Q28 - Domaines éthique",
        "Q29 - Autres éthique",
        "Q30 - Politique environnementale",
        "Q31 - Système management environnemental",
        "Q32 - Indicateurs environnementaux",
        "Q33 - Détails certification",
        "Q34 - Substances restreintes",
        "Q35 - Procédures substances",
        "Q36 - Exigences RSE fournisseurs",
        "Q37 - Communication exigences fournisseurs",
        "Q38 - Autres communication fournisseurs",
        "Q39 - Formation RSE",
        "Q40 - Personne informelle RSE",
        "Q41 - Détails contact informel",
        "Q42 - Indicateurs de base",
        "Q43 - Autres indicateurs",
        "Q44 - Règles écrites",
        "Q45 - Autres règles écrites",
        "Q46 - Intérêt accompagnement",
        "Q47 - Mesure des déchets",
        "Q48 - Réduction des déchets",
        "Q49 - Exemples réduction déchets",
        "Q50 - Filières tri/valorisation",
        "Q51 - Types déchets triés",
        "Q52 - Mesure consommation énergie",
        "Q53 - Actions réduction énergie",
        "Q54 - Exemples réduction énergie",
        "Q55 - Mesure consommation eau",
        "Q56 - Actions réduction eau",
        "Q57 - Exemples réduction eau",
        "Q58 - Actions carbone transport",
        "Q59 - Exemples transport",
        "Q60 - Mesure CO2 transport",
        "Q61 - Éco-conception",
        "Q62 - Produits éco-conçus",
        "Q63 - Commentaires",
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
}
