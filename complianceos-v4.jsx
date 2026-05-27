import { useState, useEffect, useReducer, useRef, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════
// i18n — English / German / French
// ═══════════════════════════════════════════════════
const T = {
  en: {
    appName:"ComplianceOS", dashboard:"Dashboard", checklist:"Checklist",
    team:"Team", audit:"Audit Trail", reports:"Reports", settings:"Settings",
    score:"Score", compliant:"Compliant", inProgress:"In Progress",
    nonCompliant:"Non-Compliant", notAssessed:"Not Assessed",
    highRisk:"High Risk Pending", searchPlaceholder:"Search requirements...",
    aiAdvisor:"AI Advisor", generatePolicy:"Generate Policy",
    exportExcel:"Excel Export", exportPDF:"PDF Report", exportCSV:"CSV Export",
    importCSV:"Import CSV", saveNote:"Save", addEvidence:"+ Add Evidence",
    deadline:"Deadline", notes:"Notes / Evidence", fix:"Fix",
    markDone:"Mark Done", viewAll:"View All →", noHighRisk:"All high-risk items addressed!",
    disclaimer:"Legal Disclaimer: Compliance guidance only — not legal advice.",
    notLegal:"Always consult a qualified legal professional.",
    privacyPolicy:"Privacy Policy", termsOfService:"Terms of Service",
    sort:"Sort", sortRisk:"By Risk", sortDeadline:"By Deadline",
    sortAlpha:"A-Z", sortStatus:"By Status", undo:"Undo", redo:"Redo",
    notifications:"Notifications", enableNotif:"Enable Deadline Alerts",
    shareReport:"Share Report", copied:"Copied!", postUpdate:"Post Update",
    teamActivity:"Team Activity", yourProfile:"Your Profile",
    companyName:"Company Name", yourName:"Your Name", language:"Language",
    theme:"Theme", darkMode:"Dark Mode", lightMode:"Light Mode",
    dataManagement:"Data Management", exportBackup:"Export Backup",
    resetData:"Reset All Data", integrations:"Integrations", comingSoon:"Coming Soon",
    templates:"Policy Templates", selectTemplate:"Select Template",
    download:"Download", copyText:"Copy Text", remindTeam:"Email Team Reminder",
    onboardingWelcome:"Welcome to ComplianceOS!",
    onboardingStep1:"This is your Dashboard — see your compliance score at a glance.",
    onboardingStep2:"Use Checklist to mark requirements compliant, in-progress, or non-compliant.",
    onboardingStep3:"AI Advisor gives step-by-step guidance for each requirement.",
    onboardingStep4:"Generate legal policies instantly with AI Policy Generator.",
    onboardingStep5:"Export PDF or Excel reports for auditors anytime.",
    next:"Next", prev:"Previous", finish:"Start Exploring",
    keyboardShortcuts:"Keyboard Shortcuts", shortcutSearch:"/ — Focus Search",
    shortcutUndo:"Ctrl+Z — Undo", shortcutDark:"Ctrl+D — Toggle Theme",
    shortcutAudit:"Ctrl+L — Open Audit Trail",
  },
  de: {
    appName:"ComplianceOS", dashboard:"Dashboard", checklist:"Checkliste",
    team:"Team", audit:"Prüfprotokoll", reports:"Berichte", settings:"Einstellungen",
    score:"Punktzahl", compliant:"Konform", inProgress:"In Bearbeitung",
    nonCompliant:"Nicht konform", notAssessed:"Nicht bewertet",
    highRisk:"Hohes Risiko ausstehend", searchPlaceholder:"Anforderungen suchen...",
    aiAdvisor:"KI-Berater", generatePolicy:"Richtlinie erstellen",
    exportExcel:"Excel-Export", exportPDF:"PDF-Bericht", exportCSV:"CSV-Export",
    importCSV:"CSV importieren", saveNote:"Speichern", addEvidence:"+ Nachweis hinzufügen",
    deadline:"Frist", notes:"Notizen / Nachweise", fix:"Lösung",
    markDone:"Als erledigt markieren", viewAll:"Alle anzeigen →",
    noHighRisk:"Alle Hochrisiko-Elemente erledigt!",
    disclaimer:"Rechtlicher Hinweis: Nur Compliance-Leitfaden — kein Rechtsrat.",
    notLegal:"Konsultieren Sie immer einen qualifizierten Rechtsanwalt.",
    privacyPolicy:"Datenschutzrichtlinie", termsOfService:"Nutzungsbedingungen",
    sort:"Sortieren", sortRisk:"Nach Risiko", sortDeadline:"Nach Frist",
    sortAlpha:"A-Z", sortStatus:"Nach Status", undo:"Rückgängig", redo:"Wiederholen",
    notifications:"Benachrichtigungen", enableNotif:"Fristalarmierung aktivieren",
    shareReport:"Bericht teilen", copied:"Kopiert!", postUpdate:"Update posten",
    teamActivity:"Teamaktivität", yourProfile:"Ihr Profil",
    companyName:"Unternehmensname", yourName:"Ihr Name", language:"Sprache",
    theme:"Design", darkMode:"Dunkelmodus", lightMode:"Hellmodus",
    dataManagement:"Datenverwaltung", exportBackup:"Backup exportieren",
    resetData:"Alle Daten zurücksetzen", integrations:"Integrationen",
    comingSoon:"Demnächst", templates:"Richtlinienvorlagen",
    selectTemplate:"Vorlage auswählen", download:"Herunterladen",
    copyText:"Text kopieren", remindTeam:"Team per E-Mail erinnern",
    onboardingWelcome:"Willkommen bei ComplianceOS!",
    onboardingStep1:"Dies ist Ihr Dashboard — sehen Sie Ihre Compliance-Punktzahl auf einen Blick.",
    onboardingStep2:"Nutzen Sie die Checkliste, um Anforderungen zu markieren.",
    onboardingStep3:"Der KI-Berater gibt schrittweise Anleitungen.",
    onboardingStep4:"Erstellen Sie sofort rechtliche Richtlinien mit dem KI-Generator.",
    onboardingStep5:"Exportieren Sie PDF- oder Excel-Berichte für Prüfer.",
    next:"Weiter", prev:"Zurück", finish:"Erkunden Sie",
    keyboardShortcuts:"Tastaturkürzel", shortcutSearch:"/ — Suche",
    shortcutUndo:"Strg+Z — Rückgängig", shortcutDark:"Strg+D — Design",
    shortcutAudit:"Strg+L — Prüfprotokoll",
  },
  fr: {
    appName:"ComplianceOS", dashboard:"Tableau de bord", checklist:"Liste de contrôle",
    team:"Équipe", audit:"Journal d'audit", reports:"Rapports", settings:"Paramètres",
    score:"Score", compliant:"Conforme", inProgress:"En cours",
    nonCompliant:"Non conforme", notAssessed:"Non évalué",
    highRisk:"Risque élevé en attente", searchPlaceholder:"Rechercher des exigences...",
    aiAdvisor:"Conseiller IA", generatePolicy:"Générer une politique",
    exportExcel:"Export Excel", exportPDF:"Rapport PDF", exportCSV:"Export CSV",
    importCSV:"Importer CSV", saveNote:"Enregistrer", addEvidence:"+ Ajouter preuve",
    deadline:"Échéance", notes:"Notes / Preuves", fix:"Solution",
    markDone:"Marquer comme fait", viewAll:"Voir tout →",
    noHighRisk:"Tous les éléments à risque élevé traités!",
    disclaimer:"Avis juridique: Conseils de conformité uniquement — pas de conseil juridique.",
    notLegal:"Consultez toujours un professionnel juridique qualifié.",
    privacyPolicy:"Politique de confidentialité", termsOfService:"Conditions d'utilisation",
    sort:"Trier", sortRisk:"Par risque", sortDeadline:"Par échéance",
    sortAlpha:"A-Z", sortStatus:"Par statut", undo:"Annuler", redo:"Rétablir",
    notifications:"Notifications", enableNotif:"Activer les alertes d'échéance",
    shareReport:"Partager le rapport", copied:"Copié!", postUpdate:"Publier mise à jour",
    teamActivity:"Activité d'équipe", yourProfile:"Votre profil",
    companyName:"Nom de l'entreprise", yourName:"Votre nom", language:"Langue",
    theme:"Thème", darkMode:"Mode sombre", lightMode:"Mode clair",
    dataManagement:"Gestion des données", exportBackup:"Exporter sauvegarde",
    resetData:"Réinitialiser toutes les données", integrations:"Intégrations",
    comingSoon:"Bientôt disponible", templates:"Modèles de politique",
    selectTemplate:"Sélectionner modèle", download:"Télécharger",
    copyText:"Copier le texte", remindTeam:"Rappel par e-mail",
    onboardingWelcome:"Bienvenue sur ComplianceOS!",
    onboardingStep1:"C'est votre tableau de bord — voyez votre score en un coup d'œil.",
    onboardingStep2:"Utilisez la liste de contrôle pour marquer les exigences.",
    onboardingStep3:"Le conseiller IA fournit des conseils étape par étape.",
    onboardingStep4:"Générez des politiques juridiques instantanément.",
    onboardingStep5:"Exportez des rapports PDF ou Excel pour les auditeurs.",
    next:"Suivant", prev:"Précédent", finish:"Commencer",
    keyboardShortcuts:"Raccourcis clavier", shortcutSearch:"/ — Recherche",
    shortcutUndo:"Ctrl+Z — Annuler", shortcutDark:"Ctrl+D — Thème",
    shortcutAudit:"Ctrl+L — Journal d'audit",
  }
};

// ═══════════════════════════════════════════════════
// POLICY TEMPLATES
// ═══════════════════════════════════════════════════
const POLICY_TEMPLATES = {
  "Privacy Policy (GDPR)": `PRIVACY POLICY
[COMPANY NAME] | Effective: [DATE]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. DATA CONTROLLER
[COMPANY NAME], [ADDRESS], [COUNTRY]
Contact: [EMAIL] | DPO: [DPO EMAIL]

2. DATA WE COLLECT
• Contact data: name, email, phone
• Usage data: IP address, browser type, pages visited
• Business data: company name, job title
• Payment data: processed by [PAYMENT PROCESSOR] — we do not store card details

3. PURPOSE & LEGAL BASIS (Art. 6 GDPR)
• Contract performance: order fulfilment, support
• Legitimate interest: security, fraud prevention, analytics
• Consent: marketing communications (opt-in only)
• Legal obligation: tax records, regulatory reporting

4. DATA RETENTION
• Customer data: 7 years (tax/legal obligation)
• Marketing data: until consent withdrawn
• Support data: 3 years after last contact
• Analytics: 26 months (anonymised after)

5. YOUR RIGHTS (Art. 15-22 GDPR)
You have the right to: Access · Rectification · Erasure · Restriction
· Portability · Object · Withdraw consent
Submit requests to: [EMAIL] | Response within 30 days

6. THIRD PARTIES
We share data with: [LIST PROCESSORS e.g. AWS, Stripe, Mailchimp]
All processors have signed Data Processing Agreements (DPAs).
No data sold to third parties.

7. INTERNATIONAL TRANSFERS
Data may be transferred outside the EEA under Standard Contractual Clauses (2021).

8. SECURITY
AES-256 encryption at rest · TLS 1.2+ in transit · Annual penetration testing · RBAC

9. COOKIES
See our Cookie Policy at [WEBSITE]/cookies

10. COMPLAINTS
You may lodge a complaint with your national DPA.
EU DPA directory: edpb.europa.eu

⚠️ Template only — review with qualified legal counsel before use.`,

  "Data Processing Agreement (DPA)": `DATA PROCESSING AGREEMENT
Between [DATA CONTROLLER] and [DATA PROCESSOR]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PARTIES:
Controller: [COMPANY NAME], [ADDRESS] ("Controller")
Processor: [PROCESSOR NAME], [ADDRESS] ("Processor")

1. SUBJECT MATTER & PURPOSE
The Processor processes personal data on behalf of the Controller for: [PURPOSE e.g. "providing SaaS compliance management services"]

2. NATURE OF PROCESSING
Type: [collection, storage, analysis, deletion]
Categories of data subjects: [employees, customers, users]
Categories of personal data: [contact data, usage data, financial data]
Duration: For the term of the main service agreement.

3. PROCESSOR OBLIGATIONS (Art. 28 GDPR)
The Processor shall:
(a) Process data only on documented instructions from the Controller
(b) Ensure confidentiality obligations on authorised personnel
(c) Implement appropriate technical and organisational security measures (Art. 32)
(d) Not engage sub-processors without prior written consent
(e) Assist the Controller in fulfilling data subject rights requests
(f) Delete or return all personal data on termination
(g) Provide all information necessary to demonstrate compliance
(h) Allow audits and inspections by the Controller

4. SUB-PROCESSORS
Current approved sub-processors: [LIST]
The Processor shall give 30 days notice of new sub-processors.

5. SECURITY MEASURES (Art. 32)
• Encryption at rest and in transit
• Access controls and authentication
• Regular security testing
• Business continuity planning
• Breach notification within 72 hours

6. DATA BREACH NOTIFICATION
Processor shall notify Controller without undue delay (and within 24 hours) of becoming aware of any personal data breach.

7. INTERNATIONAL TRANSFERS
Transfers outside EEA conducted under: [SCCs / Adequacy Decision / BCRs]

8. TERM AND TERMINATION
This DPA remains in force for the duration of the service agreement. On termination, Processor shall delete all Controller personal data within 30 days.

9. GOVERNING LAW
[JURISDICTION] law governs this DPA.

Signed:
Controller: _________________ Date: _______
Processor: _________________ Date: _______

⚠️ Template only — review with qualified legal counsel before use.`,

  "Cookie Policy": `COOKIE POLICY
[COMPANY NAME] | Last updated: [DATE]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. WHAT ARE COOKIES
Cookies are small text files stored on your device when you visit our website. They help us provide, improve, and personalise our services.

2. COOKIES WE USE

STRICTLY NECESSARY (Cannot be disabled)
• Session cookie: Maintains your login session
• Security cookie: CSRF protection
• Load balancer: Routes requests efficiently
Legal basis: Legitimate interest (essential for website function)

FUNCTIONAL (Optional — requires consent)  
• Language preference: Remembers your language choice
• Theme preference: Remembers dark/light mode
• Consent record: Stores your cookie choices
Legal basis: Consent

ANALYTICS (Optional — requires consent)
• [ANALYTICS TOOL e.g. Plausible/Matomo]: Page views, session duration
• We use privacy-friendly analytics with IP anonymisation
Legal basis: Consent

MARKETING (Optional — requires consent)
• [IF APPLICABLE: retargeting, ad platforms]
• Currently: We do not use marketing cookies.
Legal basis: Consent

3. MANAGING COOKIES
Change your preferences anytime: [WEBSITE]/cookie-settings
Browser settings: Most browsers allow you to block/delete cookies.
Note: Blocking necessary cookies may affect website functionality.

4. THIRD-PARTY COOKIES
[LIST any third-party services and their cookie policies]

5. RETENTION PERIODS
Session cookies: Deleted when browser closes
Persistent cookies: Maximum 12 months

6. MORE INFORMATION
Privacy Policy: [WEBSITE]/privacy
Contact: [EMAIL]

⚠️ Template only — review with qualified legal counsel before use.`,

  "Breach Response Plan": `DATA BREACH RESPONSE PLAN
[COMPANY NAME] | Version: 1.0 | Date: [DATE]
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. PURPOSE
This plan ensures [COMPANY NAME] can respond to personal data breaches promptly, minimising harm and meeting GDPR Art. 33-34 obligations (72-hour DPA notification).

2. BREACH RESPONSE TEAM
• Data Protection Officer (DPO): [NAME] | [EMAIL] | [PHONE]
• IT Security Lead: [NAME] | [EMAIL] | [PHONE]  
• Legal Counsel: [NAME] | [EMAIL] | [PHONE]
• Communications Lead: [NAME] | [EMAIL] | [PHONE]
• CEO/Management: [NAME] | [EMAIL] | [PHONE]

3. WHAT COUNTS AS A BREACH (Art. 4(12))
A breach is any accidental or unlawful destruction, loss, alteration, unauthorised disclosure of, or access to personal data. This includes:
• Ransomware / cyberattack
• Accidental deletion of data
• Email sent to wrong recipient
• Lost/stolen device containing personal data
• Unauthorised access by employee or third party

4. RESPONSE PROCEDURE

HOUR 0-4: DETECT & CONTAIN
☐ Person who discovers breach alerts DPO immediately
☐ Do NOT attempt to fix without IT involvement
☐ Preserve all logs and evidence
☐ Isolate affected systems if safe to do so
☐ Identify: What data? How many people? How long?

HOUR 4-24: ASSESS & DOCUMENT
☐ Assess likelihood and severity of harm to individuals
☐ Document: nature of breach, categories of data, approximate number of records
☐ Determine if notification to DPA is required (Art. 33)
☐ Determine if notification to individuals is required (Art. 34)
☐ Brief CEO and legal counsel

HOUR 24-72: NOTIFY DPA (if required)
☐ Submit notification to [NATIONAL DPA] at: [DPA PORTAL URL]
☐ Include: nature of breach, categories/volumes of data, likely consequences, measures taken
☐ If full information not yet available, notify with available information and supplement later

HOUR 72+: NOTIFY INDIVIDUALS (if high risk)
☐ Contact affected individuals in clear, plain language
☐ Explain: what happened, what data, what we are doing, how to protect themselves
☐ Provide contact for questions

5. NOTIFICATION TEMPLATE (DPA)
To: [NATIONAL DPA]
Subject: Personal Data Breach Notification — [COMPANY NAME]
Date of breach: [DATE] | Discovered: [DATE/TIME]
Nature: [e.g. unauthorised access to customer database]
Data categories: [e.g. names, email addresses, order history]
Approx. records: [NUMBER]
Likely consequences: [e.g. risk of phishing attacks]
Measures taken: [e.g. system secured, passwords reset, forensic investigation started]

6. POST-BREACH REVIEW
Within 14 days of resolution:
☐ Root cause analysis
☐ Update security measures to prevent recurrence
☐ Update this plan if necessary
☐ Train relevant staff
☐ Document lessons learned

7. NATIONAL DPA CONTACTS
🇩🇪 Germany: bfdi.bund.de
🇳🇱 Netherlands: autoriteitpersoonsgegevens.nl
🇫🇷 France: cnil.fr
🇬🇧 UK: ico.org.uk
🇸🇪 Sweden: imy.se
All EU DPAs: edpb.europa.eu/about-edpb/board/members_en

⚠️ Template only — review with qualified legal counsel before use.`
};

// ═══════════════════════════════════════════════════
// REQUIREMENTS DATA
// ═══════════════════════════════════════════════════
const REGS = {
  GDPR:[
    {id:"G1",cat:"Data Collection",title:"Lawful Basis Documentation",desc:"Document lawful basis for each processing activity",risk:"HIGH",ref:"Art. 6",fix:"Create ROPA documenting legal basis per data type"},
    {id:"G2",cat:"Data Collection",title:"Privacy Notice / Policy",desc:"Plain-language privacy notice publicly accessible",risk:"HIGH",ref:"Art. 13-14",fix:"Publish notice covering data types, purposes, retention, rights"},
    {id:"G3",cat:"Data Collection",title:"Special Category Data Safeguards",desc:"Extra protection for health, biometric, political data",risk:"HIGH",ref:"Art. 9",fix:"Identify special category data; obtain explicit consent"},
    {id:"G4",cat:"Data Collection",title:"Children's Data Protection",desc:"Parental consent for under-16",risk:"HIGH",ref:"Art. 8",fix:"Add age verification; define threshold per country"},
    {id:"G5",cat:"Consent",title:"Cookie Consent Banner (CMP)",desc:"Explicit opt-in before placing non-essential cookies",risk:"HIGH",ref:"Art. 7 + ePrivacy",fix:"Deploy CMP — no pre-ticked boxes; granular controls"},
    {id:"G6",cat:"Consent",title:"Consent Records & Logs",desc:"Prove when and how consent was obtained",risk:"MEDIUM",ref:"Art. 7(1)",fix:"Log timestamp, IP hash, consent version, choices"},
    {id:"G7",cat:"Consent",title:"Consent Withdrawal Mechanism",desc:"Withdrawal as easy as giving consent",risk:"HIGH",ref:"Art. 7(3)",fix:"'Manage Preferences' link in footer; one-click withdrawal"},
    {id:"G8",cat:"Consent",title:"Marketing Opt-in",desc:"Explicit opt-in for email/SMS marketing",risk:"MEDIUM",ref:"Art. 7 + ePrivacy",fix:"Remove pre-checked boxes; double opt-in; unsubscribe in all comms"},
    {id:"G9",cat:"Rights",title:"Right to Access (DSAR)",desc:"Respond within 1 month, free of charge",risk:"HIGH",ref:"Art. 15",fix:"Create DSAR form; assign owner; 30-day calendar reminder"},
    {id:"G10",cat:"Rights",title:"Right to Erasure",desc:"Delete all personal data upon valid request",risk:"HIGH",ref:"Art. 17",fix:"Map all data stores; build deletion workflow; document refusals"},
    {id:"G11",cat:"Rights",title:"Right to Portability",desc:"Provide data in machine-readable format",risk:"MEDIUM",ref:"Art. 20",fix:"Build export feature for user personal data"},
    {id:"G12",cat:"Rights",title:"Right to Rectification",desc:"Correct inaccurate data within 1 month",risk:"MEDIUM",ref:"Art. 16",fix:"Create correction process; propagate to all systems"},
    {id:"G13",cat:"Rights",title:"Right to Restriction",desc:"Restrict processing in specific circumstances",risk:"LOW",ref:"Art. 18",fix:"Implement restriction flag; ensure flagged records not processed"},
    {id:"G14",cat:"Rights",title:"Right to Object",desc:"Allow objection to legitimate interest & marketing",risk:"MEDIUM",ref:"Art. 21",fix:"Opt-out for profiling/marketing; mandatory stop on objection"},
    {id:"G15",cat:"Security",title:"Encryption at Rest & Transit",desc:"AES-256 stored; TLS 1.2+ transmitted",risk:"HIGH",ref:"Art. 32",fix:"Enable DB encryption; enforce HTTPS+HSTS; encrypt storage"},
    {id:"G16",cat:"Security",title:"Role-Based Access Control",desc:"Least privilege — access only what needed",risk:"HIGH",ref:"Art. 32",fix:"Audit accounts; implement RBAC; disable inactive after 30 days; MFA"},
    {id:"G17",cat:"Security",title:"Breach Notification (72h)",desc:"Notify DPA within 72 hours of breach",risk:"HIGH",ref:"Art. 33-34",fix:"Write incident response plan; assign DPO; create notification template"},
    {id:"G18",cat:"Security",title:"Penetration Testing",desc:"Regular security testing of data systems",risk:"MEDIUM",ref:"Art. 32",fix:"Annual pen test; quarterly scans; patch critical CVEs in 30 days"},
    {id:"G19",cat:"Security",title:"Backup & Disaster Recovery",desc:"Tested backup procedures",risk:"MEDIUM",ref:"Art. 32",fix:"3-2-1 backup; test quarterly; document RTO/RPO"},
    {id:"G20",cat:"Processors",title:"Data Processing Agreements (DPAs)",desc:"Signed DPAs with ALL third-party processors",risk:"HIGH",ref:"Art. 28",fix:"Audit all vendors; request/sign DPAs before sharing data"},
    {id:"G21",cat:"Processors",title:"Sub-processor Management",desc:"Approved sub-processor list; 30-day notice",risk:"HIGH",ref:"Art. 28(4)",fix:"Publish sub-processor register; notification clause in DPAs"},
    {id:"G22",cat:"Processors",title:"Vendor Security Assessment",desc:"Annual security review of processors",risk:"MEDIUM",ref:"Art. 28(3)(h)",fix:"Vendor questionnaire; request ISO 27001/SOC2; audit rights"},
    {id:"G23",cat:"Retention",title:"Data Retention Policy",desc:"Defined retention periods per data category",risk:"MEDIUM",ref:"Art. 5(1)(e)",fix:"Create retention schedule; implement automated deletion"},
    {id:"G24",cat:"Retention",title:"Data Minimization",desc:"Collect only strictly necessary data",risk:"MEDIUM",ref:"Art. 5(1)(c)",fix:"Audit forms/APIs; remove unnecessary fields"},
    {id:"G25",cat:"Retention",title:"Automated Data Deletion",desc:"Technical enforcement of retention periods",risk:"MEDIUM",ref:"Art. 5(1)(e)",fix:"Scheduled deletion/anonymization jobs; test quarterly"},
    {id:"G26",cat:"Accountability",title:"Data Protection Officer (DPO)",desc:"Appoint DPO if large-scale processing",risk:"MEDIUM",ref:"Art. 37",fix:"Assess DPO requirement; appoint internally or outsource"},
    {id:"G27",cat:"Accountability",title:"GDPR Staff Training",desc:"Annual training for all staff handling data",risk:"MEDIUM",ref:"Art. 39",fix:"Mandatory training; maintain records; update annually"},
    {id:"G28",cat:"Accountability",title:"ROPA (Record of Processing Activities)",desc:"Mandatory register of all processing activities",risk:"HIGH",ref:"Art. 30",fix:"Document per activity: purpose, categories, recipients, retention"},
    {id:"G29",cat:"Accountability",title:"Privacy by Design & Default",desc:"Embed privacy into product architecture",risk:"MEDIUM",ref:"Art. 25",fix:"Privacy review in sprints; default to most private settings"},
    {id:"G30",cat:"Accountability",title:"DPIA (Data Protection Impact Assessment)",desc:"Mandatory before high-risk processing",risk:"HIGH",ref:"Art. 35",fix:"Identify high-risk activities; conduct DPIA before launch"},
    {id:"G31",cat:"Transfers",title:"International Data Flow Mapping",desc:"Map all data flows outside EEA",risk:"HIGH",ref:"Art. 44-46",fix:"Audit all flows; document legal basis per transfer"},
    {id:"G32",cat:"Transfers",title:"Standard Contractual Clauses (SCCs)",desc:"2021 EU SCCs for non-adequate countries",risk:"HIGH",ref:"Art. 46(2)(c)",fix:"Sign updated 2021 SCCs with all non-EEA processors"},
    {id:"G33",cat:"Transfers",title:"Transfer Impact Assessment (TIA)",desc:"Assess receiving country protection level",risk:"MEDIUM",ref:"Art. 46",fix:"Document TIA; assess govt access risks; supplementary measures"},
  ],
  CCPA:[
    {id:"C1",cat:"Consumer Rights",title:"Right to Know",desc:"Disclose PI categories collected in last 12 months",risk:"HIGH",ref:"§1798.100",fix:"Update privacy policy with all PI categories, sources, purposes"},
    {id:"C2",cat:"Consumer Rights",title:"Right to Delete",desc:"Delete consumer PI within 45 days of request",risk:"HIGH",ref:"§1798.105",fix:"Build deletion form; 45-day response; propagate to service providers"},
    {id:"C3",cat:"Consumer Rights",title:"Right to Opt-Out of Sale",desc:"'Do Not Sell or Share My PI' link required",risk:"HIGH",ref:"§1798.120",fix:"Add DNSMI link site-wide; honor within 15 business days"},
    {id:"C4",cat:"Consumer Rights",title:"Right to Correct",desc:"Correct inaccurate PI upon verified request",risk:"MEDIUM",ref:"§1798.106",fix:"Correction request process; 45-day response; notify service providers"},
    {id:"C5",cat:"Consumer Rights",title:"Right to Limit Sensitive PI",desc:"'Limit Use of My Sensitive PI' opt-out required",risk:"HIGH",ref:"§1798.121",fix:"Add LSPI link in footer; restrict sensitive PI to necessary purposes"},
    {id:"C6",cat:"Business",title:"California Privacy Policy",desc:"Privacy policy with CPRA-specific disclosures",risk:"HIGH",ref:"§1798.130",fix:"Add CA section: PI categories, rights, opt-out links, retention"},
    {id:"C7",cat:"Business",title:"Identity Verification Process",desc:"Verify consumer identity before fulfilling requests",risk:"MEDIUM",ref:"§1798.130(a)",fix:"2-factor verification; document method; avoid over-collecting"},
    {id:"C8",cat:"Business",title:"Non-Discrimination Policy",desc:"No discrimination for exercising CCPA rights",risk:"MEDIUM",ref:"§1798.125",fix:"Document policy; review pricing tiers; train support staff"},
  ],
  "AI ACT":[
    {id:"A1",cat:"Classification",title:"AI System Risk Assessment",desc:"Classify all AI: Prohibited / High-Risk / Limited / Minimal",risk:"HIGH",ref:"Art. 6-7",fix:"Inventory AI systems; apply Annex III criteria; document classification"},
    {id:"A2",cat:"Prohibited",title:"Prohibited AI Practices Audit",desc:"No social scoring, real-time biometrics, manipulation",risk:"HIGH",ref:"Art. 5",fix:"Audit AI against Art. 5 prohibited list; halt prohibited systems"},
    {id:"A3",cat:"High-Risk",title:"High-Risk AI Conformity",desc:"Technical docs, logs, human oversight required",risk:"HIGH",ref:"Art. 9-15",fix:"Document AI systems; implement logging; mandatory human review"},
    {id:"A4",cat:"Transparency",title:"AI Transparency & Disclosure",desc:"Disclose AI interaction; label AI-generated content",risk:"HIGH",ref:"Art. 50",fix:"Add disclosure banner; label AI content; watermarking"},
    {id:"A5",cat:"Governance",title:"AI Governance Framework",desc:"Internal AI policies, ethics review, accountability",risk:"MEDIUM",ref:"Art. 9",fix:"Create AI governance policy; appoint AI compliance officer"},
    {id:"A6",cat:"Governance",title:"Technical Documentation & Logs",desc:"Tech docs + logs — 6 month minimum retention",risk:"MEDIUM",ref:"Art. 11-12",fix:"Document AI purpose, data, accuracy; logging with 6-month retention"},
  ],
  "ISO 27001":[
    {id:"I1",cat:"Policies",title:"Information Security Policy",desc:"Documented, approved, communicated security policy",risk:"HIGH",ref:"Clause 5.2",fix:"Draft and board-approve policy; communicate to all staff; annual review"},
    {id:"I2",cat:"Risk",title:"Information Security Risk Assessment",desc:"Systematic risk identification and evaluation",risk:"HIGH",ref:"Clause 6.1.2",fix:"Conduct risk assessment; document threats/vulnerabilities; risk register"},
    {id:"I3",cat:"Risk",title:"Risk Treatment Plan",desc:"Documented plan for treating identified risks",risk:"HIGH",ref:"Clause 6.1.3",fix:"Select controls; document treatment plan; assign owners; target dates"},
    {id:"I4",cat:"Access",title:"Access Management Controls",desc:"User registration, provisioning, and review",risk:"HIGH",ref:"Annex A.9",fix:"Joiners/movers/leavers process; quarterly access reviews; PAM"},
    {id:"I5",cat:"Cryptography",title:"Cryptography Policy",desc:"Policy governing cryptographic controls",risk:"MEDIUM",ref:"Annex A.10",fix:"Define approved algorithms; document key management; annual review"},
    {id:"I6",cat:"Suppliers",title:"Supplier Security Management",desc:"Security requirements in supplier agreements",risk:"MEDIUM",ref:"Annex A.15",fix:"Security clauses in contracts; assess critical suppliers; audit rights"},
    {id:"I7",cat:"Incidents",title:"Security Incident Management",desc:"Documented incident response procedures",risk:"HIGH",ref:"Annex A.16",fix:"Create incident response plan; log all incidents; post-incident reviews"},
    {id:"I8",cat:"Continuity",title:"Business Continuity Management",desc:"Tested continuity plans for critical systems",risk:"MEDIUM",ref:"Annex A.17",fix:"Identify critical systems; document BCPs; test annually"},
  ],
  HIPAA:[
    {id:"H1",cat:"Privacy",title:"Privacy Rule Compliance",desc:"Policies for PHI use and disclosure",risk:"HIGH",ref:"45 CFR §164.500",fix:"Implement privacy policies; appoint Privacy Officer; staff training"},
    {id:"H2",cat:"Security",title:"Administrative Safeguards",desc:"Security management, workforce training, contingency",risk:"HIGH",ref:"45 CFR §164.308",fix:"Risk analysis; security management; train workforce; contingency plan"},
    {id:"H3",cat:"Security",title:"Physical Safeguards",desc:"Facility access controls and workstation security",risk:"MEDIUM",ref:"45 CFR §164.310",fix:"Facility access controls; workstation policies; device disposal"},
    {id:"H4",cat:"Security",title:"Technical Safeguards",desc:"Access controls, audit controls, transmission security",risk:"HIGH",ref:"45 CFR §164.312",fix:"Unique user IDs; automatic logoff; encryption; audit controls"},
    {id:"H5",cat:"Breach",title:"Breach Notification Rule",desc:"Notify HHS and individuals within 60 days",risk:"HIGH",ref:"45 CFR §164.400",fix:"Define breach; response procedure; notify within 60 days"},
    {id:"H6",cat:"Agreements",title:"Business Associate Agreements (BAAs)",desc:"Signed BAAs with all entities accessing PHI",risk:"HIGH",ref:"45 CFR §164.314",fix:"Identify business associates; execute BAAs; include required provisions"},
  ]
};

const REG_COLOR={GDPR:"#00e5a0",CCPA:"#00b4d8","AI ACT":"#a855f7","ISO 27001":"#f59e0b",HIPAA:"#f43f5e"};
const rc=r=>r==="HIGH"?"#ff4d4f":r==="MEDIUM"?"#faad14":"#52c41a";
const rb=r=>r==="HIGH"?"rgba(255,77,79,0.1)":r==="MEDIUM"?"rgba(250,173,20,0.1)":"rgba(82,196,26,0.1)";
const fmtDate=d=>d?new Date(d).toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"}):"—";
const fmtTs=ts=>new Date(ts).toLocaleString("en-GB",{day:"numeric",month:"short",hour:"2-digit",minute:"2-digit"});

// Storage
const sGet=async(key,def=null,shared=false)=>{try{if(window.storage){const r=await window.storage.get(key,shared);return r?JSON.parse(r.value):def;}const v=localStorage.getItem("cos4_"+key);return v?JSON.parse(v):def;}catch{return def;}};
const sSet=async(key,val,shared=false)=>{try{if(window.storage)await window.storage.set(key,JSON.stringify(val),shared);else localStorage.setItem("cos4_"+key,JSON.stringify(val));}catch{}};

// Undo/Redo reducer
function undoReducer(state,action){
  switch(action.type){
    case "SET":{
      const next=[...state.past,state.present].slice(-20);
      return{past:next,present:action.payload,future:[]};
    }
    case "UNDO":{
      if(!state.past.length)return state;
      const prev=state.past[state.past.length-1];
      return{past:state.past.slice(0,-1),present:prev,future:[state.present,...state.future]};
    }
    case "REDO":{
      if(!state.future.length)return state;
      const next=state.future[0];
      return{past:[...state.past,state.present],present:next,future:state.future.slice(1)};
    }
    default:return state;
  }
}

// ═══════════════════════════════════════════════════
// DISCLAIMER MODAL
// ═══════════════════════════════════════════════════
function DisclaimerModal({onAccept}){
  const [ok,setOk]=useState(false);
  return(
    <div style={{position:"fixed",inset:0,background:"#050912",zIndex:9999,display:"flex",alignItems:"center",justifyContent:"center",padding:20,overflow:"auto"}}>
      <div style={{maxWidth:600,width:"100%"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:52,height:52,borderRadius:14,background:"linear-gradient(135deg,#00e5a0,#00b4d8)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,fontWeight:900,color:"#000",margin:"0 auto 14px"}}>C</div>
          <h1 style={{fontSize:26,fontWeight:900,fontFamily:"'Syne',sans-serif",color:"#fff",marginBottom:6}}>ComplianceOS v4.0</h1>
          <p style={{color:"rgba(255,255,255,0.4)",fontSize:13}}>GDPR · CCPA · AI Act · ISO 27001 · HIPAA · 58 Requirements</p>
        </div>
        <div style={{background:"rgba(255,77,79,0.07)",border:"1px solid rgba(255,77,79,0.2)",borderRadius:12,padding:"16px 20px",marginBottom:14}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:6,color:"#ff6b6b"}}>⚠️ Legal Disclaimer</div>
          <p style={{fontSize:13,color:"rgba(255,255,255,0.6)",lineHeight:1.7}}>This tool provides <strong style={{color:"#fff"}}>compliance guidance only</strong> — NOT legal advice. Scores are indicative. ComplianceOS accepts <strong style={{color:"#fff"}}>no liability</strong> for fines or penalties. Always consult a qualified legal professional.</p>
        </div>
        <div style={{background:"rgba(255,255,255,0.03)",border:"1px solid rgba(255,255,255,0.07)",borderRadius:12,padding:"16px 20px",marginBottom:14}}>
          {["58 GDPR/CCPA/AI Act/ISO 27001/HIPAA requirements","AI Advisor + AI Policy Generator","Multi-language: English, Deutsch, Français","PWA installable · Offline-capable · Dark & Light mode","Team collaboration · Audit trail · CSV/Excel/PDF export","Browser deadline notifications · Undo/Redo · Keyboard shortcuts"].map((f,i)=>(
            <div key={i} style={{display:"flex",gap:10,alignItems:"flex-start",marginBottom:i<5?7:0}}>
              <span style={{color:"#00e5a0",flexShrink:0,fontSize:12}}>✓</span>
              <span style={{fontSize:12,color:"rgba(255,255,255,0.55)"}}>{f}</span>
            </div>
          ))}
        </div>
        <label style={{display:"flex",gap:12,alignItems:"flex-start",cursor:"pointer",marginBottom:16,padding:"12px 16px",background:"rgba(0,229,160,0.05)",border:"1px solid rgba(0,229,160,0.15)",borderRadius:10}}>
          <input type="checkbox" checked={ok} onChange={e=>setOk(e.target.checked)} style={{width:16,height:16,accentColor:"#00e5a0",flexShrink:0,marginTop:2}}/>
          <span style={{fontSize:12,color:"rgba(255,255,255,0.6)",lineHeight:1.6}}>I understand this tool does not provide legal advice. I accept the Terms of Service and Privacy Policy.</span>
        </label>
        <button disabled={!ok} onClick={onAccept} style={{width:"100%",padding:"14px",borderRadius:11,border:"none",cursor:ok?"pointer":"not-allowed",background:ok?"linear-gradient(135deg,#00e5a0,#00b4d8)":"rgba(255,255,255,0.07)",color:ok?"#000":"rgba(255,255,255,0.2)",fontSize:15,fontWeight:800,fontFamily:"'Syne',sans-serif",transition:"all 0.2s"}}>
          {ok?"Enter ComplianceOS →":"Accept disclaimer to continue"}
        </button>
      </div>
      <style>{`@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}@keyframes dot{0%,100%{opacity:.3;transform:scale(.8)}50%{opacity:1;transform:scale(1.2)}}`}</style>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// ONBOARDING TOUR
// ═══════════════════════════════════════════════════
function OnboardingTour({lang,onFinish}){
  const t=T[lang];
  const [step,setStep]=useState(0);
  const steps=[
    {title:t.onboardingWelcome,desc:t.onboardingStep1,icon:"◈"},
    {title:"Checklist",desc:t.onboardingStep2,icon:"✓"},
    {title:"AI Advisor",desc:t.onboardingStep3,icon:"✦"},
    {title:"Policy Generator",desc:t.onboardingStep4,icon:"⚡"},
    {title:"Reports",desc:t.onboardingStep5,icon:"↗"},
  ];
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",backdropFilter:"blur(10px)",zIndex:3000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"linear-gradient(135deg,#0f1923,#0a1628)",border:"1px solid rgba(0,229,160,0.25)",borderRadius:20,padding:"36px",maxWidth:480,width:"100%",textAlign:"center",boxShadow:"0 40px 80px rgba(0,0,0,0.6)"}}>
        <div style={{width:64,height:64,borderRadius:18,background:"rgba(0,229,160,0.1)",border:"1px solid rgba(0,229,160,0.2)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,color:"#00e5a0",margin:"0 auto 20px"}}>{steps[step].icon}</div>
        <h3 style={{fontSize:20,fontWeight:900,fontFamily:"'Syne',sans-serif",marginBottom:12,color:"#fff"}}>{steps[step].title}</h3>
        <p style={{fontSize:14,color:"rgba(255,255,255,0.6)",lineHeight:1.7,marginBottom:28}}>{steps[step].desc}</p>
        <div style={{display:"flex",gap:6,justifyContent:"center",marginBottom:24}}>
          {steps.map((_,i)=>(
            <div key={i} style={{width:i===step?20:7,height:7,borderRadius:4,background:i===step?"#00e5a0":"rgba(255,255,255,0.15)",transition:"all 0.3s ease"}}/>
          ))}
        </div>
        <div style={{display:"flex",gap:10}}>
          {step>0&&<button onClick={()=>setStep(s=>s-1)} style={{flex:1,padding:"11px",borderRadius:9,border:"1px solid rgba(255,255,255,0.1)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.6)",cursor:"pointer",fontSize:13}}>{t.prev}</button>}
          <button onClick={()=>step<steps.length-1?setStep(s=>s+1):onFinish()} style={{flex:2,padding:"11px",borderRadius:9,border:"none",background:"linear-gradient(135deg,#00e5a0,#00b4d8)",color:"#000",cursor:"pointer",fontSize:13,fontWeight:800,fontFamily:"'Syne',sans-serif"}}>
            {step<steps.length-1?t.next:t.finish}
          </button>
        </div>
        <button onClick={onFinish} style={{background:"none",border:"none",color:"rgba(255,255,255,0.25)",fontSize:11,cursor:"pointer",marginTop:14}}>Skip tour</button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// AI ADVISOR
// ═══════════════════════════════════════════════════
function AIAdvisor({item,reg,lang,onClose}){
  const t=T[lang];
  const [loading,setLoading]=useState(true);
  const [text,setText]=useState("");
  const [err,setErr]=useState("");
  useEffect(()=>{
    (async()=>{
      try{
        const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:`You are a senior compliance lawyer. Advise a European SMB on: "${item.title}" (${item.ref}) | ${reg} | Risk: ${item.risk}\n\nProvide:\n1. WHY this matters legally (2 sentences, mention real fines)\n2. ACTION PLAN: 4 numbered steps — practical, specific\n3. Time to implement\n4. One free tool/resource\n\nUnder 280 words. End: "⚠️ AI guidance only — verify with qualified legal counsel."`}]})});
        const d=await res.json();setText(d.content?.[0]?.text||"No response.");
      }catch{setErr("AI service unavailable. Please try again.");}
      finally{setLoading(false);}
    })();
  },[]);
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.8)",backdropFilter:"blur(8px)",zIndex:1500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"linear-gradient(135deg,#0f1923,#0a1628)",border:"1px solid rgba(0,229,160,0.2)",borderRadius:18,padding:"28px",maxWidth:560,width:"100%",maxHeight:"85vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
          <div>
            <div style={{background:"rgba(0,229,160,0.1)",border:"1px solid rgba(0,229,160,0.2)",borderRadius:5,padding:"2px 9px",display:"inline-block",marginBottom:8}}>
              <span style={{fontSize:10,color:"#00e5a0",fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{t.aiAdvisor.toUpperCase()}</span>
            </div>
            <h3 style={{color:"#fff",fontSize:15,fontFamily:"'Syne',sans-serif",fontWeight:800,marginBottom:3}}>{item.title}</h3>
            <span style={{color:"rgba(255,255,255,0.35)",fontSize:11}}>{item.ref} · {reg}</span>
          </div>
          <button onClick={onClose} style={{background:"rgba(255,255,255,0.08)",border:"none",color:"#fff",width:32,height:32,borderRadius:9,cursor:"pointer",fontSize:16,flexShrink:0}}>×</button>
        </div>
        <div style={{background:"rgba(255,77,79,0.06)",border:"1px solid rgba(255,77,79,0.12)",borderRadius:8,padding:"7px 12px",marginBottom:12,fontSize:11,color:"rgba(255,140,140,0.8)"}}>⚠️ AI guidance — not legal advice</div>
        <div style={{flex:1,overflow:"auto",background:"rgba(255,255,255,0.02)",borderRadius:11,padding:16,minHeight:140}}>
          {loading?(<div style={{display:"flex",flexDirection:"column",gap:10,alignItems:"center",justifyContent:"center",height:140}}><div style={{display:"flex",gap:6}}>{[0,1,2].map(i=><div key={i} style={{width:7,height:7,borderRadius:"50%",background:"#00e5a0",animation:`dot 1.2s ease-in-out ${i*0.2}s infinite`}}/>)}</div><span style={{color:"rgba(255,255,255,0.35)",fontSize:12}}>Analysing...</span></div>)
          :err?<p style={{color:"#ff4d4f",fontSize:12}}>{err}</p>
          :<pre style={{color:"rgba(255,255,255,0.82)",fontSize:12,lineHeight:1.8,whiteSpace:"pre-wrap",fontFamily:"'Space Grotesk',sans-serif"}}>{text}</pre>}
        </div>
        <style>{`@keyframes dot{0%,100%{opacity:.3;transform:scale(.8)}50%{opacity:1;transform:scale(1.2)}}`}</style>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// POLICY GENERATOR
// ═══════════════════════════════════════════════════
function PolicyGenerator({company,lang,onClose}){
  const t=T[lang];
  const [mode,setMode]=useState("ai");
  const [type,setType]=useState("Privacy Policy (GDPR)");
  const [country,setCountry]=useState("Germany");
  const [industry,setIndustry]=useState("Technology");
  const [loading,setLoading]=useState(false);
  const [result,setResult]=useState("");
  const [generated,setGenerated]=useState(false);
  const [selectedTemplate,setSelectedTemplate]=useState(Object.keys(POLICY_TEMPLATES)[0]);

  const generate=async()=>{
    setLoading(true);setGenerated(false);setResult("");
    try{
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:`Generate a professional ${type} for: Company: ${company}, Country: ${country}, Industry: ${industry}. GDPR-compliant. Use [COMPANY NAME], [EMAIL], [ADDRESS] as placeholders. Include disclaimer that template needs legal review.`}]})});
      const d=await res.json();
      setResult(d.content?.[0]?.text||"Generation failed.");setGenerated(true);
    }catch{setResult("AI service unavailable.");setGenerated(true);}
    finally{setLoading(false);}
  };

  const dl=(content,name)=>{const b=new Blob([content],{type:"text/plain"});const u=URL.createObjectURL(b);const a=document.createElement("a");a.href=u;a.download=name;a.click();};

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",backdropFilter:"blur(10px)",zIndex:1500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"linear-gradient(135deg,#0f1923,#0a1628)",border:"1px solid rgba(168,85,247,0.25)",borderRadius:20,padding:"28px",maxWidth:680,width:"100%",maxHeight:"90vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}>
          <div>
            <div style={{background:"rgba(168,85,247,0.1)",border:"1px solid rgba(168,85,247,0.2)",borderRadius:5,padding:"2px 9px",display:"inline-block",marginBottom:8}}>
              <span style={{fontSize:10,color:"#a855f7",fontFamily:"'DM Mono',monospace",letterSpacing:1}}>POLICY GENERATOR</span>
            </div>
            <h3 style={{color:"#fff",fontSize:17,fontFamily:"'Syne',sans-serif",fontWeight:900}}>{t.generatePolicy}</h3>
          </div>
          <button onClick={onClose} style={{background:"rgba(255,255,255,0.08)",border:"none",color:"#fff",width:32,height:32,borderRadius:9,cursor:"pointer",fontSize:16,flexShrink:0}}>×</button>
        </div>
        <div style={{display:"flex",gap:8,marginBottom:18}}>
          {["ai","templates"].map(m=>(
            <button key={m} onClick={()=>setMode(m)} style={{flex:1,padding:"8px",borderRadius:8,border:"1px solid",borderColor:mode===m?"rgba(168,85,247,0.4)":"rgba(255,255,255,0.08)",background:mode===m?"rgba(168,85,247,0.1)":"rgba(255,255,255,0.03)",color:mode===m?"#a855f7":"rgba(255,255,255,0.4)",cursor:"pointer",fontSize:12,fontWeight:600}}>
              {m==="ai"?"✦ AI Generator":"📄 Template Library"}
            </button>
          ))}
        </div>
        {mode==="ai"?(
          !generated?(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                {[["Document Type",["Privacy Policy (GDPR)","Cookie Policy","Data Processing Agreement (DPA)","Employee Data Policy","Data Breach Response Plan","Data Retention Policy"],type,setType],
                  ["Jurisdiction",["Germany","France","Netherlands","UK","Sweden","Poland","Spain","Italy","Austria","Belgium"],country,setCountry],
                  ["Industry",["Technology","Healthcare","Finance","E-commerce","Education","Legal Services","Manufacturing","Consulting","Media"],industry,setIndustry]
                ].map(([label,opts,val,set],i)=>(
                  <div key={i}>
                    <label style={{fontSize:10,color:"rgba(255,255,255,0.3)",display:"block",marginBottom:5,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{label.toUpperCase()}</label>
                    <select value={val} onChange={e=>set(e.target.value)} style={{width:"100%",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:7,color:"#fff",padding:"8px 10px",fontSize:12,outline:"none"}}>
                      {opts.map(o=><option key={o} value={o} style={{background:"#0f1923"}}>{o}</option>)}
                    </select>
                  </div>
                ))}
              </div>
              <div style={{background:"rgba(255,77,79,0.05)",border:"1px solid rgba(255,77,79,0.12)",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:11,color:"rgba(255,140,140,0.8)"}}>⚠️ AI template — mandatory legal review before use</div>
              <button onClick={generate} disabled={loading} style={{width:"100%",padding:"12px",borderRadius:10,border:"none",cursor:"pointer",background:"linear-gradient(135deg,#a855f7,#7c3aed)",color:"#fff",fontSize:14,fontWeight:800,fontFamily:"'Syne',sans-serif"}}>
                {loading?"Generating...":"Generate Document →"}
              </button>
              {loading&&<div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:7,marginTop:12,color:"rgba(255,255,255,0.35)",fontSize:12}}><div style={{display:"flex",gap:5}}>{[0,1,2].map(i=><div key={i} style={{width:5,height:5,borderRadius:"50%",background:"#a855f7",animation:`dot 1.2s ease-in-out ${i*0.2}s infinite`}}/>)}</div>Generating...</div>}
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",flex:1,overflow:"hidden"}}>
              <div style={{flex:1,overflow:"auto",background:"rgba(255,255,255,0.02)",borderRadius:11,padding:16,marginBottom:12}}>
                <pre style={{fontSize:12,color:"rgba(255,255,255,0.75)",lineHeight:1.8,whiteSpace:"pre-wrap",fontFamily:"'Space Grotesk',sans-serif"}}>{result}</pre>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>dl(result,`${type.replace(/\s+/g,"-")}.txt`)} style={{flex:1,padding:"10px",borderRadius:8,border:"1px solid rgba(168,85,247,0.3)",background:"rgba(168,85,247,0.1)",color:"#a855f7",cursor:"pointer",fontSize:12,fontWeight:700}}>↓ {t.download}</button>
                <button onClick={()=>navigator.clipboard?.writeText(result)} style={{padding:"10px 16px",borderRadius:8,border:"1px solid rgba(255,255,255,0.1)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.5)",cursor:"pointer",fontSize:12}}>{t.copyText}</button>
                <button onClick={()=>setGenerated(false)} style={{padding:"10px 16px",borderRadius:8,border:"1px solid rgba(255,255,255,0.1)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.5)",cursor:"pointer",fontSize:12}}>← New</button>
              </div>
            </div>
          )
        ):(
          <div style={{display:"flex",flexDirection:"column",flex:1,overflow:"hidden"}}>
            <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
              {Object.keys(POLICY_TEMPLATES).map(k=>(
                <button key={k} onClick={()=>setSelectedTemplate(k)} style={{padding:"5px 12px",borderRadius:16,border:"1px solid",borderColor:selectedTemplate===k?"rgba(168,85,247,0.4)":"rgba(255,255,255,0.08)",background:selectedTemplate===k?"rgba(168,85,247,0.1)":"rgba(255,255,255,0.03)",color:selectedTemplate===k?"#a855f7":"rgba(255,255,255,0.4)",cursor:"pointer",fontSize:11,whiteSpace:"nowrap"}}>
                  {k.split(" ").slice(0,3).join(" ")}
                </button>
              ))}
            </div>
            <div style={{flex:1,overflow:"auto",background:"rgba(255,255,255,0.02)",borderRadius:11,padding:16,marginBottom:12}}>
              <pre style={{fontSize:12,color:"rgba(255,255,255,0.7)",lineHeight:1.8,whiteSpace:"pre-wrap",fontFamily:"'Space Grotesk',sans-serif"}}>{POLICY_TEMPLATES[selectedTemplate]}</pre>
            </div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>dl(POLICY_TEMPLATES[selectedTemplate],`${selectedTemplate.replace(/[\s()]/g,"-")}.txt`)} style={{flex:1,padding:"10px",borderRadius:8,border:"1px solid rgba(168,85,247,0.3)",background:"rgba(168,85,247,0.1)",color:"#a855f7",cursor:"pointer",fontSize:12,fontWeight:700}}>↓ {t.download}</button>
              <button onClick={()=>navigator.clipboard?.writeText(POLICY_TEMPLATES[selectedTemplate])} style={{padding:"10px 16px",borderRadius:8,border:"1px solid rgba(255,255,255,0.1)",background:"rgba(255,255,255,0.05)",color:"rgba(255,255,255,0.5)",cursor:"pointer",fontSize:12}}>{t.copyText}</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function LegalModal({title,content,onClose}){
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.85)",backdropFilter:"blur(10px)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#0c1520",border:"1px solid rgba(255,255,255,0.1)",borderRadius:18,padding:"28px",maxWidth:660,width:"100%",maxHeight:"82vh",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <h2 style={{fontFamily:"'Syne',sans-serif",fontSize:17,fontWeight:800}}>{title}</h2>
          <button onClick={onClose} style={{background:"rgba(255,255,255,0.08)",border:"none",color:"#fff",width:32,height:32,borderRadius:9,cursor:"pointer",fontSize:16}}>×</button>
        </div>
        <pre style={{overflow:"auto",fontSize:12,color:"rgba(255,255,255,0.6)",lineHeight:1.8,whiteSpace:"pre-wrap",fontFamily:"'DM Mono',monospace",flex:1}}>{content}</pre>
        <button onClick={onClose} style={{marginTop:16,padding:"10px",borderRadius:9,border:"1px solid rgba(0,229,160,0.3)",background:"rgba(0,229,160,0.08)",color:"#00e5a0",cursor:"pointer",fontSize:13}}>Close</button>
      </div>
    </div>
  );
}

function ScoreGauge({score,size=118}){
  const r=46,c=2*Math.PI*r,off=c-(score/100)*c;
  const color=score>=80?"#00e5a0":score>=50?"#faad14":"#ff4d4f";
  return(
    <div style={{position:"relative",width:size,height:size,flexShrink:0}}>
      <svg width={size} height={size} style={{transform:"rotate(-90deg)"}}>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke="rgba(255,255,255,0.05)" strokeWidth={9}/>
        <circle cx={size/2} cy={size/2} r={r} fill="none" stroke={color} strokeWidth={9} strokeDasharray={c} strokeDashoffset={off} strokeLinecap="round" style={{transition:"stroke-dashoffset 1s ease,stroke 0.4s"}}/>
      </svg>
      <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
        <span style={{fontSize:size*.26,fontWeight:900,color,fontFamily:"'Syne',sans-serif",lineHeight:1}}>{score}</span>
        <span style={{fontSize:9,color:"rgba(255,255,255,0.3)",letterSpacing:2,fontFamily:"'DM Mono',monospace"}}>SCORE</span>
      </div>
    </div>
  );
}

function CookieBanner({t,onAccept}){
  return(
    <div style={{position:"fixed",bottom:16,left:"50%",transform:"translateX(-50%)",zIndex:900,width:"calc(100% - 32px)",maxWidth:660,background:"#0f1923",border:"1px solid rgba(255,255,255,0.1)",borderRadius:12,padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap",boxShadow:"0 16px 50px rgba(0,0,0,0.5)"}}>
      <div style={{display:"flex",gap:10,alignItems:"center"}}>
        <span style={{fontSize:18}}>🍪</span>
        <div>
          <div style={{fontSize:12,fontWeight:600,marginBottom:2}}>Functional Storage Only</div>
          <div style={{fontSize:11,color:"rgba(255,255,255,0.4)"}}>No tracking or advertising. Only functional data storage.</div>
        </div>
      </div>
      <button onClick={onAccept} style={{padding:"7px 18px",borderRadius:7,border:"none",background:"#00e5a0",color:"#000",fontSize:12,fontWeight:700,cursor:"pointer",flexShrink:0}}>Accept</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════
export default function ComplianceOS(){
  const [init,setInit]=useState(false);
  const [accepted,setAccepted]=useState(false);
  const [cookieOk,setCookieOk]=useState(false);
  const [showTour,setShowTour]=useState(false);
  const [lang,setLang]=useState("en");
  const [theme,setTheme]=useState("dark");
  const [statusHistory,dispatch]=useReducer(undoReducer,{past:[],present:{},future:[]});
  const [deadlines,setDeadlines]=useState({});
  const [notes,setNotes]=useState({});
  const [evidence,setEvidence]=useState({});
  const [auditLog,setAuditLog]=useState([]);
  const [teamFeed,setTeamFeed]=useState([]);
  const [companyName,setCompanyName]=useState("Your Company GmbH");
  const [userName,setUserName]=useState("Compliance Officer");
  const [activeReg,setActiveReg]=useState("GDPR");
  const [activeTab,setActiveTab]=useState("dashboard");
  const [aiItem,setAiItem]=useState(null);
  const [showPolicy,setShowPolicy]=useState(false);
  const [showPrivacy,setShowPrivacy]=useState(false);
  const [showTerms,setShowTerms]=useState(false);
  const [catFilter,setCatFilter]=useState("All");
  const [riskFilter,setRiskFilter]=useState("All");
  const [statusFilter,setStatusFilter]=useState("All");
  const [search,setSearch]=useState("");
  const [sortBy,setSortBy]=useState("risk");
  const [expanded,setExpanded]=useState(null);
  const [editNote,setEditNote]=useState(null);
  const [tempNote,setTempNote]=useState("");
  const [newEvidence,setNewEvidence]=useState(null);
  const [tempEvidence,setTempEvidence]=useState("");
  const [teamMsg,setTeamMsg]=useState("");
  const [notifEnabled,setNotifEnabled]=useState(false);
  const [shareMsg,setShareMsg]=useState("");
  const [editName,setEditName]=useState(false);
  const [mobile,setMobile]=useState(false);
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [undoMsg,setUndoMsg]=useState("");
  const [csvImportMsg,setCsvImportMsg]=useState("");
  const [regUpdateResult,setRegUpdateResult]=useState("");
  const [checkingUpdates,setCheckingUpdates]=useState(false);
  const searchRef=useRef(null);
  const t=T[lang];
  const statuses=statusHistory.present;

  useEffect(()=>{
    const check=()=>setMobile(window.innerWidth<768);
    check();window.addEventListener("resize",check);return()=>window.removeEventListener("resize",check);
  },[]);

  // Load data
  useEffect(()=>{
    (async()=>{
      const [acc,ck,st,dl,nt,ev,al,tf,cn,un,lg,th,tour]=await Promise.all([
        sGet("accepted",false),sGet("cookies",false),sGet("statuses",{}),
        sGet("deadlines",{}),sGet("notes",{}),sGet("evidence",{}),
        sGet("audit",[]),sGet("teamfeed",[],true),
        sGet("company","Your Company GmbH"),sGet("username","Compliance Officer"),
        sGet("lang","en"),sGet("theme","dark"),sGet("tourDone",false)
      ]);
      setAccepted(acc);setCookieOk(ck);
      dispatch({type:"SET",payload:st});
      setDeadlines(dl);setNotes(nt);setEvidence(ev);setAuditLog(al);
      setTeamFeed(tf);setCompanyName(cn);setUserName(un);
      setLang(lg);setTheme(th);
      if(acc&&!tour)setShowTour(true);
      setInit(true);
    })();
  },[]);

  // Theme CSS
  useEffect(()=>{
    const root=document.documentElement;
    if(theme==="light"){
      root.style.setProperty("--bg","#f0f4f8");
      root.style.setProperty("--surface","#ffffff");
      root.style.setProperty("--surface2","#f7fafc");
      root.style.setProperty("--border","rgba(0,0,0,0.08)");
      root.style.setProperty("--text","#1a202c");
      root.style.setProperty("--text2","#4a5568");
      root.style.setProperty("--text3","#718096");
    }else{
      root.style.setProperty("--bg","#080e18");
      root.style.setProperty("--surface","rgba(255,255,255,0.025)");
      root.style.setProperty("--surface2","rgba(255,255,255,0.02)");
      root.style.setProperty("--border","rgba(255,255,255,0.06)");
      root.style.setProperty("--text","#ffffff");
      root.style.setProperty("--text2","rgba(255,255,255,0.6)");
      root.style.setProperty("--text3","rgba(255,255,255,0.3)");
    }
  },[theme]);

  // Keyboard shortcuts
  useEffect(()=>{
    const handler=e=>{
      if(e.key==="/"&&document.activeElement!==searchRef.current){e.preventDefault();setActiveTab("checklist");setTimeout(()=>searchRef.current?.focus(),100);}
      if((e.ctrlKey||e.metaKey)&&e.key==="z"){e.preventDefault();handleUndo();}
      if((e.ctrlKey||e.metaKey)&&e.key==="d"){e.preventDefault();toggleTheme();}
      if((e.ctrlKey||e.metaKey)&&e.key==="l"){e.preventDefault();setActiveTab("audit");}
      if(e.key==="Escape"){setAiItem(null);setShowPolicy(false);setSidebarOpen(false);}
    };
    window.addEventListener("keydown",handler);
    return()=>window.removeEventListener("keydown",handler);
  },[statuses,statusHistory]);

  // Deadline notifications check
  useEffect(()=>{
    if(!notifEnabled)return;
    const check=()=>{
      const today=new Date();
      Object.entries(deadlines).forEach(([id,date])=>{
        if(!date)return;
        const dl=new Date(date);
        const diff=Math.ceil((dl-today)/(1000*60*60*24));
        if(diff<=7&&diff>=0&&statuses[id]!=="compliant"){
          const req=Object.values(REGS).flat().find(r=>r.id===id);
          if(req&&Notification.permission==="granted"){
            new Notification("ComplianceOS Deadline Alert",{body:`"${req.title}" deadline in ${diff} day(s)!`,icon:"/favicon.ico"});
          }
        }
      });
    };
    check();
    const interval=setInterval(check,3600000);
    return()=>clearInterval(interval);
  },[notifEnabled,deadlines,statuses]);

  const addAudit=useCallback(async(action,item,detail="")=>{
    const entry={id:Date.now(),action,item,detail,ts:new Date().toISOString(),user:userName};
    const next=[entry,...auditLog.slice(0,199)];
    setAuditLog(next);await sSet("audit",next);
  },[auditLog,userName]);

  const addTeamFeed=useCallback(async(msg)=>{
    const entry={id:Date.now(),msg,user:userName,company:companyName,ts:new Date().toISOString()};
    const next=[entry,...teamFeed.slice(0,49)];
    setTeamFeed(next);await sSet("teamfeed",next,true);
  },[teamFeed,userName,companyName]);

  const toggle=useCallback(async(id,val,title)=>{
    const cur=statuses[id],nxt=cur===val?undefined:val;
    const next={...statuses,[id]:nxt};
    dispatch({type:"SET",payload:next});
    await sSet("statuses",next);
    const action=nxt==="compliant"?"Marked Compliant":nxt==="non-compliant"?"Marked Non-Compliant":nxt==="in-progress"?"Marked In Progress":"Status Cleared";
    await addAudit(action,title);
    await addTeamFeed(`${action}: ${title}`);
  },[statuses,addAudit,addTeamFeed]);

  const handleUndo=useCallback(async()=>{
    if(!statusHistory.past.length)return;
    dispatch({type:"UNDO"});
    await sSet("statuses",statusHistory.past[statusHistory.past.length-1]||{});
    setUndoMsg(t.undo);setTimeout(()=>setUndoMsg(""),2000);
  },[statusHistory,t]);

  const handleRedo=useCallback(async()=>{
    if(!statusHistory.future.length)return;
    dispatch({type:"REDO"});
    await sSet("statuses",statusHistory.future[0]||{});
  },[statusHistory]);

  const toggleTheme=()=>{const n=theme==="dark"?"light":"dark";setTheme(n);sSet("theme",n);};
  const saveNote=async(id,title)=>{const n={...notes,[id]:tempNote};setNotes(n);await sSet("notes",n);await addAudit("Note Updated",title);setEditNote(null);};
  const saveDeadline=async(id,date,title)=>{const n={...deadlines,[id]:date};setDeadlines(n);await sSet("deadlines",n);await addAudit("Deadline Set",title,date?fmtDate(date):"Cleared");};
  const addEvidenceItem=async(id,title)=>{if(!tempEvidence.trim())return;const cur=evidence[id]||[];const n={...evidence,[id]:[{text:tempEvidence,ts:new Date().toISOString()},...cur.slice(0,4)]};setEvidence(n);await sSet("evidence",n);await addAudit("Evidence Added",title,tempEvidence.slice(0,40));setNewEvidence(null);setTempEvidence("");};
  const acceptDisclaimer=async()=>{await sSet("accepted",true);setAccepted(true);setShowTour(true);};
  const acceptCookie=async()=>{await sSet("cookies",true);setCookieOk(true);};
  const finishTour=async()=>{setShowTour(false);await sSet("tourDone",true);};

  const enableNotifications=async()=>{
    if(!("Notification"in window)){alert("Notifications not supported in this browser.");return;}
    const perm=await Notification.requestPermission();
    if(perm==="granted"){setNotifEnabled(true);new Notification("ComplianceOS",{body:"Deadline alerts enabled! You'll be notified 7 days before deadlines."});}
    else alert("Please enable notifications in browser settings.");
  };

  const shareReport=()=>{
    const text=`ComplianceOS Report — ${companyName}\nRegulation: ${activeReg}\nScore: ${score}%\nCompliant: ${compliant}/${total}\nGenerated: ${new Date().toLocaleDateString()}\n\nView full app at: complianceos.vercel.app`;
    navigator.clipboard?.writeText(text).then(()=>{setShareMsg(t.copied);setTimeout(()=>setShareMsg(""),2000);});
  };

  const exportCSV=()=>{
    const reqs=REGS[activeReg]||[];
    const rows=[["ID","Category","Requirement","Risk","Reference","Status","Deadline","Notes","Fix"],...reqs.map(r=>[r.id,r.cat,r.title,r.risk,r.ref,statuses[r.id]||"Not Assessed",deadlines[r.id]?fmtDate(deadlines[r.id]):"",notes[r.id]||"",r.fix])];
    const csv=rows.map(r=>r.map(c=>`"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
    const b=new Blob([csv],{type:"text/csv"});const u=URL.createObjectURL(b);const a=document.createElement("a");a.href=u;a.download=`${companyName.replace(/\s+/g,"-")}-${activeReg}-compliance.csv`;a.click();
    addAudit("CSV Export",activeReg);
  };

  const importCSV=async(file)=>{
    if(!file)return;
    const text=await file.text();
    const lines=text.split("\n").slice(1);
    let count=0;
    const newStatuses={...statuses};
    lines.forEach(line=>{
      if(!line.trim())return;
      const cols=line.match(/(".*?"|[^,]+)/g)||[];
      const id=cols[0]?.replace(/"/g,"").trim();
      const status=cols[5]?.replace(/"/g,"").trim().toLowerCase().replace(/\s+/g,"-");
      if(id&&["compliant","non-compliant","in-progress"].includes(status)){newStatuses[id]=status;count++;}
    });
    dispatch({type:"SET",payload:newStatuses});
    await sSet("statuses",newStatuses);
    await addAudit("CSV Import",`${count} statuses updated`);
    setCsvImportMsg(`✓ ${count} statuses imported`);setTimeout(()=>setCsvImportMsg(""),3000);
  };

  const exportExcel=()=>{
    const reqs=REGS[activeReg]||[];
    const rows=reqs.map(r=>({Regulation:activeReg,ID:r.id,Category:r.cat,Requirement:r.title,Risk:r.risk,Reference:r.ref,Status:statuses[r.id]||"Not Assessed",Deadline:deadlines[r.id]?fmtDate(deadlines[r.id]):"",Notes:notes[r.id]||"","Evidence Count":(evidence[r.id]||[]).length,Fix:r.fix}));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,activeReg);
    const ar=auditLog.map(a=>({Timestamp:fmtTs(a.ts),User:a.user||"",Action:a.action,Requirement:a.item,Detail:a.detail||""}));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(ar),"Audit Log");
    XLSX.writeFile(wb,`${companyName.replace(/\s+/g,"-")}-${activeReg}-${new Date().toISOString().split("T")[0]}.xlsx`);
    addAudit("Excel Export",activeReg);
  };

  const exportPDF=()=>{
    const reqs=REGS[activeReg]||[];
    const html=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${companyName} ${activeReg}</title><style>body{font-family:Arial,sans-serif;color:#111;padding:32px;font-size:12px}h1{font-size:22px;font-weight:900}h2{font-size:15px;margin:22px 0 10px;border-bottom:2px solid #e5e7eb;padding-bottom:5px}.hdr{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:28px;padding-bottom:18px;border-bottom:3px solid #000}.score{font-size:48px;font-weight:900;color:${score>=80?"#16a34a":score>=50?"#d97706":"#dc2626"}}.badge{display:inline-block;padding:2px 7px;border-radius:4px;font-size:11px;font-weight:700}.HIGH{background:#fee2e2;color:#dc2626}.MEDIUM{background:#fef3c7;color:#d97706}.LOW{background:#dcfce7;color:#16a34a}.compliant{background:#dcfce7;color:#16a34a}.non-compliant{background:#fee2e2;color:#dc2626}.in-progress{background:#fef3c7;color:#d97706}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:22px}.stat{background:#f9fafb;border-radius:6px;padding:12px;text-align:center}.stat-n{font-size:24px;font-weight:900}.stat-l{font-size:10px;color:#6b7280;margin-top:3px}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#f3f4f6;padding:7px 8px;text-align:left;font-size:10px;text-transform:uppercase}td{padding:7px 8px;border-bottom:1px solid #e5e7eb;vertical-align:top}.disc{background:#fef3c7;border:1px solid #d97706;border-radius:6px;padding:12px;margin-top:30px;font-size:11px}@media print{.disc{break-inside:avoid}}</style></head><body><div class="hdr"><div><h1>${companyName}</h1><p style="color:#6b7280;margin:3px 0">${activeReg} Compliance Report — ${new Date().toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"})}</p><p style="color:#6b7280;font-size:10px">Prepared by ${userName} · ComplianceOS v4.0</p></div><div style="text-align:right"><div class="score">${score}%</div><div style="font-size:11px;color:#6b7280">Overall Score</div></div></div><h2>Summary</h2><div class="stats"><div class="stat"><div class="stat-n" style="color:#16a34a">${compliant}</div><div class="stat-l">Compliant</div></div><div class="stat"><div class="stat-n" style="color:#d97706">${inProgress}</div><div class="stat-l">In Progress</div></div><div class="stat"><div class="stat-n" style="color:#dc2626">${nonCompliant}</div><div class="stat-l">Non-Compliant</div></div><div class="stat"><div class="stat-n" style="color:#6b7280">${reqs.length-compliant-inProgress-nonCompliant}</div><div class="stat-l">Not Assessed</div></div></div><h2>Requirements</h2><table><thead><tr><th>Ref</th><th>Requirement</th><th>Risk</th><th>Status</th><th>Deadline</th><th>Notes</th></tr></thead><tbody>${reqs.map(r=>`<tr><td style="font-family:monospace;font-size:10px;white-space:nowrap">${r.ref}</td><td><strong>${r.title}</strong><br><span style="color:#6b7280;font-size:10px">${r.cat}</span></td><td><span class="badge ${r.risk}">${r.risk}</span></td><td><span class="badge ${statuses[r.id]||"not-assessed"}">${statuses[r.id]||"Not Assessed"}</span></td><td style="font-size:10px;white-space:nowrap">${deadlines[r.id]?fmtDate(deadlines[r.id]):"—"}</td><td style="font-size:10px;color:#6b7280">${notes[r.id]||"—"}</td></tr>`).join("")}</tbody></table><div class="disc"><strong>⚠️ LEGAL DISCLAIMER</strong> — This report is for informational purposes only. It does NOT constitute legal advice. Always consult a qualified legal professional. ComplianceOS accepts no liability for fines or penalties.</div><p style="text-align:center;font-size:10px;color:#9ca3af;margin-top:16px">ComplianceOS v4.0 · ${companyName} · ${userName} · ${new Date().toLocaleDateString("en-GB")} · Confidential</p></body></html>`;
    const w=window.open("","_blank","width=1000,height=800");w.document.write(html);w.document.close();w.onload=()=>w.print();
    addAudit("PDF Export",activeReg);
  };

  const checkRegUpdates=async()=>{
    setCheckingUpdates(true);setRegUpdateResult("");
    try{
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:600,messages:[{role:"user",content:"Summarize the most important GDPR, CCPA, and EU AI Act regulatory updates from 2024-2025. Bullet points, max 10 items. Include impact level (HIGH/MEDIUM/LOW)."}]})});
      const d=await res.json();setRegUpdateResult(d.content?.[0]?.text||"No updates found.");
    }catch{setRegUpdateResult("Unable to fetch. Check your connection.");}
    finally{setCheckingUpdates(false);}
  };

  const reqs=REGS[activeReg]||[];
  const compliant=useMemo(()=>reqs.filter(r=>statuses[r.id]==="compliant").length,[statuses,reqs]);
  const nonCompliant=useMemo(()=>reqs.filter(r=>statuses[r.id]==="non-compliant").length,[statuses,reqs]);
  const inProgress=useMemo(()=>reqs.filter(r=>statuses[r.id]==="in-progress").length,[statuses,reqs]);
  const total=reqs.length;
  const score=total?Math.round((compliant/total)*100):0;
  const highPending=reqs.filter(r=>r.risk==="HIGH"&&statuses[r.id]!=="compliant").length;

  const filtered=useMemo(()=>{
    let res=reqs.filter(r=>
      (catFilter==="All"||r.cat===catFilter)&&
      (riskFilter==="All"||r.risk===riskFilter)&&
      (statusFilter==="All"||(statusFilter==="not-assessed"?!statuses[r.id]:statuses[r.id]===statusFilter))&&
      (!search||r.title.toLowerCase().includes(search.toLowerCase())||r.desc.toLowerCase().includes(search.toLowerCase()))
    );
    if(sortBy==="risk"){const o={HIGH:0,MEDIUM:1,LOW:2};res=[...res].sort((a,b)=>o[a.risk]-o[b.risk]);}
    else if(sortBy==="alpha")res=[...res].sort((a,b)=>a.title.localeCompare(b.title));
    else if(sortBy==="deadline")res=[...res].sort((a,b)=>{const da=deadlines[a.id]?new Date(deadlines[a.id]):new Date("9999");const db=deadlines[b.id]?new Date(deadlines[b.id]):new Date("9999");return da-db;});
    else if(sortBy==="status"){const o={"non-compliant":0,"in-progress":1,undefined:2,"compliant":3};res=[...res].sort((a,b)=>(o[statuses[a.id]]??2)-(o[statuses[b.id]]??2));}
    return res;
  },[reqs,catFilter,riskFilter,statusFilter,search,sortBy,statuses,deadlines]);

  const cats=useMemo(()=>["All",...new Set(reqs.map(r=>r.cat))],[reqs]);
  const catStats=useMemo(()=>[...new Set(reqs.map(r=>r.cat))].map(cat=>{
    const items=reqs.filter(r=>r.cat===cat);
    const done=items.filter(r=>statuses[r.id]==="compliant").length;
    return{cat,done,total:items.length,pct:Math.round((done/items.length)*100)};
  }),[reqs,statuses]);

  const isDark=theme==="dark";
  const fg=isDark?"#fff":"#1a202c";
  const bg=isDark?"#080e18":"#f0f4f8";
  const surface=isDark?"rgba(255,255,255,0.025)":"#ffffff";
  const border=isDark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.08)";
  const text2=isDark?"rgba(255,255,255,0.55)":"#4a5568";
  const text3=isDark?"rgba(255,255,255,0.3)":"#718096";

  if(!init)return(
    <div style={{minHeight:"100vh",background:bg,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{textAlign:"center",color:text3}}>
        <div style={{display:"flex",gap:7,justifyContent:"center",marginBottom:12}}>{[0,1,2].map(i=><div key={i} style={{width:8,height:8,borderRadius:"50%",background:"#00e5a0",animation:`dot 1.2s ease-in-out ${i*0.2}s infinite`}}/>)}</div>
        <span style={{fontSize:13,fontFamily:"'Space Grotesk',sans-serif"}}>Loading ComplianceOS v4.0...</span>
      </div>
      <style>{`@keyframes dot{0%,100%{opacity:.3;transform:scale(.8)}50%{opacity:1;transform:scale(1.2)}}`}</style>
    </div>
  );
  if(!accepted)return <DisclaimerModal onAccept={acceptDisclaimer}/>;

  const NAV=[
    {id:"dashboard",label:t.dashboard,icon:"◈"},
    {id:"checklist",label:t.checklist,icon:"✓"},
    {id:"team",label:t.team,icon:"◎"},
    {id:"audit",label:t.audit,icon:"⊙"},
    {id:"reports",label:t.reports,icon:"↗"},
    {id:"settings",label:t.settings,icon:"⚙"},
  ];

  return(
    <div style={{minHeight:"100vh",background:bg,fontFamily:"'Space Grotesk',sans-serif",color:fg,transition:"background 0.3s,color 0.3s"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800;900&family=Space+Grotesk:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:rgba(0,229,160,0.3);border-radius:3px}
        @keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
        @keyframes dot{0%,100%{opacity:.3;transform:scale(.8)}50%{opacity:1;transform:scale(1.2)}}
        @keyframes slideIn{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
        select option{background:#0f1923;color:#fff}
        input[type=date]{color-scheme:dark}
        .nh:hover{background:rgba(255,255,255,0.06)!important}
        .rh:hover{background:rgba(0,229,160,0.03)!important}
        .bl:hover{opacity:0.8}
      `}</style>

      {/* Undo notification */}
      {undoMsg&&<div style={{position:"fixed",top:80,left:"50%",transform:"translateX(-50%)",zIndex:999,background:"#00e5a0",color:"#000",padding:"8px 20px",borderRadius:8,fontSize:13,fontWeight:700,animation:"slideIn 0.3s ease"}}>{undoMsg}</div>}
      {csvImportMsg&&<div style={{position:"fixed",top:80,left:"50%",transform:"translateX(-50%)",zIndex:999,background:"#00e5a0",color:"#000",padding:"8px 20px",borderRadius:8,fontSize:13,fontWeight:700,animation:"slideIn 0.3s ease"}}>{csvImportMsg}</div>}
      {shareMsg&&<div style={{position:"fixed",top:80,left:"50%",transform:"translateX(-50%)",zIndex:999,background:"#00e5a0",color:"#000",padding:"8px 20px",borderRadius:8,fontSize:13,fontWeight:700}}>{shareMsg}</div>}

      {/* Mobile overlay */}
      {mobile&&sidebarOpen&&<div onClick={()=>setSidebarOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:199}}/>}

      {/* Sidebar */}
      <div style={{position:"fixed",left:mobile?sidebarOpen?0:-220:0,top:0,bottom:0,width:210,background:isDark?"linear-gradient(180deg,#0c1520,#080e18)":"#fff",borderRight:`1px solid ${border}`,display:"flex",flexDirection:"column",zIndex:mobile?200:100,padding:"22px 0",transition:"left 0.25s ease",boxShadow:isDark?"none":"0 0 20px rgba(0,0,0,0.06)"}}>
        <div style={{padding:"0 16px",marginBottom:20,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:9}}>
            <div style={{width:30,height:30,borderRadius:8,background:"linear-gradient(135deg,#00e5a0,#00b4d8)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:900,color:"#000",flexShrink:0}}>C</div>
            <div>
              <div style={{fontSize:13,fontWeight:900,fontFamily:"'Syne',sans-serif",color:fg}}>ComplianceOS</div>
              <div style={{fontSize:9,color:text3,letterSpacing:2}}>v4.0</div>
            </div>
          </div>
          {mobile&&<button onClick={()=>setSidebarOpen(false)} style={{background:"none",border:"none",color:text3,fontSize:18,cursor:"pointer"}}>×</button>}
        </div>

        {/* Regulation switcher */}
        <div style={{padding:"0 9px",marginBottom:16}}>
          <div style={{fontSize:9,color:text3,letterSpacing:2,marginBottom:6,paddingLeft:5,fontFamily:"'DM Mono',monospace"}}>REGULATION</div>
          {Object.keys(REGS).map(reg=>(
            <button key={reg} onClick={()=>{setActiveReg(reg);setCatFilter("All");setSearch("");if(mobile)setSidebarOpen(false);}} className="nh" style={{width:"100%",display:"flex",alignItems:"center",justifyContent:"space-between",padding:"7px 9px",borderRadius:7,border:"none",cursor:"pointer",background:activeReg===reg?`${REG_COLOR[reg]}15`:"transparent",color:activeReg===reg?REG_COLOR[reg]:text3,fontSize:11,fontFamily:"'DM Mono',monospace",fontWeight:700,textAlign:"left",marginBottom:2,borderLeft:activeReg===reg?`2px solid ${REG_COLOR[reg]}`:"2px solid transparent",transition:"all 0.15s",letterSpacing:.5}}>
              <span>{reg}</span><span style={{fontSize:10,opacity:0.6}}>{REGS[reg].length}</span>
            </button>
          ))}
        </div>

        <div style={{height:1,background:border,margin:"0 9px 14px"}}/>

        <div style={{padding:"0 9px",flex:1}}>
          {NAV.map(item=>(
            <button key={item.id} onClick={()=>{setActiveTab(item.id);if(mobile)setSidebarOpen(false);}} className="nh" style={{width:"100%",display:"flex",alignItems:"center",gap:9,padding:"9px 9px",borderRadius:8,border:"none",cursor:"pointer",background:activeTab===item.id?"rgba(0,229,160,0.1)":"transparent",color:activeTab===item.id?"#00e5a0":text3,fontSize:13,fontWeight:500,textAlign:"left",marginBottom:2,borderLeft:activeTab===item.id?"2px solid #00e5a0":"2px solid transparent",transition:"all 0.12s"}}>
              <span style={{fontSize:14,width:16,textAlign:"center"}}>{item.icon}</span>{item.label}
            </button>
          ))}
        </div>

        <div style={{padding:"12px 10px 0"}}>
          <div onClick={()=>{setActiveTab("checklist");if(mobile)setSidebarOpen(false);}} style={{background:"rgba(255,77,79,0.07)",border:"1px solid rgba(255,77,79,0.12)",borderRadius:9,padding:"9px 11px",cursor:"pointer"}}>
            <div style={{fontSize:9,color:"#ff4d4f",fontFamily:"'DM Mono',monospace",letterSpacing:1,marginBottom:2}}>{t.highRisk.toUpperCase()}</div>
            <div style={{fontSize:22,fontWeight:900,fontFamily:"'Syne',sans-serif",color:"#ff4d4f",lineHeight:1}}>{highPending}</div>
            <div style={{fontSize:10,color:text3,marginTop:2}}>tap → fix</div>
          </div>
        </div>

        {/* Undo/Redo */}
        <div style={{padding:"10px 10px 0",display:"flex",gap:6}}>
          <button onClick={handleUndo} disabled={!statusHistory.past.length} style={{flex:1,padding:"6px",borderRadius:7,border:`1px solid ${border}`,background:"transparent",color:statusHistory.past.length?text2:text3,cursor:statusHistory.past.length?"pointer":"not-allowed",fontSize:11,transition:"all 0.15s"}} title={t.shortcutUndo}>{t.undo} ↩</button>
          <button onClick={handleRedo} disabled={!statusHistory.future.length} style={{flex:1,padding:"6px",borderRadius:7,border:`1px solid ${border}`,background:"transparent",color:statusHistory.future.length?text2:text3,cursor:statusHistory.future.length?"pointer":"not-allowed",fontSize:11,transition:"all 0.15s"}}>{t.redo} ↪</button>
        </div>

        <div style={{padding:"10px 12px 0",display:"flex",gap:8,flexWrap:"wrap"}}>
          <button onClick={()=>setShowPrivacy(true)} style={{background:"none",border:"none",color:text3,fontSize:10,cursor:"pointer"}}>{t.privacyPolicy}</button>
          <button onClick={()=>setShowTerms(true)} style={{background:"none",border:"none",color:text3,fontSize:10,cursor:"pointer"}}>{t.termsOfService}</button>
        </div>
      </div>

      {/* Main */}
      <div style={{marginLeft:mobile?0:210,minHeight:"100vh"}}>
        {/* Header */}
        <div style={{padding:mobile?"12px 14px":"14px 26px",borderBottom:`1px solid ${border}`,display:"flex",alignItems:"center",justifyContent:"space-between",background:isDark?"rgba(8,14,24,0.95)":"rgba(255,255,255,0.95)",backdropFilter:"blur(20px)",position:"sticky",top:0,zIndex:50,gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            {mobile&&<button onClick={()=>setSidebarOpen(true)} style={{background:surface,border:`1px solid ${border}`,color:fg,width:32,height:32,borderRadius:8,cursor:"pointer",fontSize:15,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>☰</button>}
            <div>
              {editName?(<input value={companyName} onChange={e=>setCompanyName(e.target.value)} onBlur={async()=>{setEditName(false);await sSet("company",companyName);}} onKeyDown={e=>e.key==="Enter"&&(setEditName(false),sSet("company",companyName))} autoFocus style={{background:"transparent",border:"none",borderBottom:`1px solid #00e5a0`,color:fg,fontSize:mobile?15:17,fontFamily:"'Syne',sans-serif",fontWeight:900,outline:"none",maxWidth:240}}/>
              ):(<h1 onClick={()=>setEditName(true)} style={{fontSize:mobile?15:17,fontWeight:900,fontFamily:"'Syne',sans-serif",cursor:"pointer",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:mobile?160:280,color:fg}} title="Click to edit">{companyName} <span style={{fontSize:11,color:text3}}>✎</span></h1>)}
              <p style={{color:text3,fontSize:10,marginTop:1}}>{activeReg} · {total} requirements · ComplianceOS v4.0</p>
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:7}}>
            {!mobile&&<button onClick={()=>setShowPolicy(true)} style={{padding:"6px 14px",borderRadius:8,border:"1px solid rgba(168,85,247,0.3)",background:"rgba(168,85,247,0.1)",color:"#a855f7",cursor:"pointer",fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>✦ {t.generatePolicy}</button>}
            {/* Language */}
            <select value={lang} onChange={e=>{setLang(e.target.value);sSet("lang",e.target.value);}} style={{background:surface,border:`1px solid ${border}`,borderRadius:7,color:fg,padding:"5px 8px",fontSize:11,outline:"none",cursor:"pointer",fontFamily:"'DM Mono',monospace"}}>
              <option value="en">🇬🇧 EN</option>
              <option value="de">🇩🇪 DE</option>
              <option value="fr">🇫🇷 FR</option>
            </select>
            {/* Theme toggle */}
            <button onClick={toggleTheme} title={t.shortcutDark} style={{background:surface,border:`1px solid ${border}`,color:fg,width:32,height:32,borderRadius:8,cursor:"pointer",fontSize:15,display:"flex",alignItems:"center",justifyContent:"center"}}>
              {isDark?"☀️":"🌙"}
            </button>
            <div style={{display:"flex",alignItems:"center",gap:5,background:surface,border:`1px solid ${border}`,borderRadius:7,padding:"4px 9px",fontSize:11,color:REG_COLOR[activeReg]}}>
              <span style={{width:5,height:5,borderRadius:"50%",background:REG_COLOR[activeReg],display:"inline-block"}}/>
              {!mobile&&activeReg}
            </div>
          </div>
        </div>

        <div style={{padding:mobile?"14px":"22px 26px",animation:"fadeIn 0.3s ease"}}>
          {/* Disclaimer banner */}
          <div style={{background:"rgba(250,173,20,0.07)",border:"1px solid rgba(250,173,20,0.17)",borderRadius:9,padding:"7px 13px",marginBottom:14,display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:13,flexShrink:0}}>⚠️</span>
            <span style={{fontSize:11,color:"rgba(255,210,80,0.8)"}}><strong>{t.disclaimer}</strong> {t.notLegal}</span>
          </div>

          {/* ───── DASHBOARD ───── */}
          {activeTab==="dashboard"&&(
            <div>
              <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(4,1fr)",gap:12,marginBottom:18}}>
                {[{label:t.score,val:`${score}%`,sub:`${compliant}/${total}`,color:"#00e5a0"},{label:t.compliant,val:compliant,sub:"done",color:"#00e5a0"},{label:t.inProgress,val:inProgress,sub:"being fixed",color:"#faad14"},{label:t.nonCompliant,val:nonCompliant,sub:"urgent",color:"#ff4d4f"}].map((s,i)=>(
                  <div key={i} style={{background:surface,border:`1px solid ${border}`,borderRadius:12,padding:"15px 17px"}}>
                    <div style={{fontSize:9,color:text3,letterSpacing:1,marginBottom:6,fontFamily:"'DM Mono',monospace"}}>{s.label.toUpperCase()}</div>
                    <div style={{fontSize:mobile?26:32,fontWeight:900,fontFamily:"'Syne',sans-serif",color:s.color,lineHeight:1}}>{s.val}</div>
                    <div style={{fontSize:11,color:text3,marginTop:4}}>{s.sub}</div>
                  </div>
                ))}
              </div>

              <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"190px 1fr",gap:15,marginBottom:15}}>
                <div style={{background:"rgba(0,229,160,0.04)",border:"1px solid rgba(0,229,160,0.1)",borderRadius:14,padding:"22px 17px",display:"flex",flexDirection:"column",alignItems:"center",gap:12}}>
                  <ScoreGauge score={score}/>
                  <div style={{textAlign:"center"}}>
                    <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",color:fg}}>{score>=80?"Strong":score>=50?"Moderate Risk":"High Risk"}</div>
                    <div style={{fontSize:11,color:text3,marginTop:3}}>Fine: <span style={{color:score>=80?"#00e5a0":"#ff4d4f",fontWeight:700}}>{score>=80?"Low":"€20M+"}</span></div>
                  </div>
                </div>
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:14,padding:"18px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:14,color:fg}}>By Category</div>
                  <div style={{display:"flex",flexDirection:"column",gap:10}}>
                    {catStats.map(({cat,done,total:t2,pct})=>(
                      <div key={cat}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                          <span style={{fontSize:12,color:text2}}>{cat}</span>
                          <span style={{fontSize:11,fontFamily:"'DM Mono',monospace",color:pct>=80?"#00e5a0":pct>=50?"#faad14":"#ff4d4f"}}>{done}/{t2}</span>
                        </div>
                        <div style={{height:4,background:isDark?"rgba(255,255,255,0.05)":"rgba(0,0,0,0.08)",borderRadius:2,overflow:"hidden"}}>
                          <div style={{height:"100%",borderRadius:2,width:`${pct}%`,background:pct>=80?"linear-gradient(90deg,#00e5a0,#00b4d8)":pct>=50?"#faad14":"#ff4d4f",transition:"width 0.8s ease"}}/>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              <div style={{background:surface,border:`1px solid ${border}`,borderRadius:14,padding:"17px 20px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:13}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",color:fg}}>🔴 {t.highRisk}</div>
                  <button onClick={()=>setActiveTab("checklist")} style={{background:"rgba(0,229,160,0.1)",border:"1px solid rgba(0,229,160,0.2)",color:"#00e5a0",padding:"4px 12px",borderRadius:6,cursor:"pointer",fontSize:11}}>{t.viewAll}</button>
                </div>
                {reqs.filter(r=>r.risk==="HIGH"&&statuses[r.id]!=="compliant").slice(0,5).map(r=>(
                  <div key={r.id} className="rh" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${border}`,gap:8,transition:"background 0.15s",borderRadius:4}}>
                    <div style={{display:"flex",alignItems:"center",gap:9,flex:1,minWidth:0}}>
                      <div style={{width:6,height:6,borderRadius:"50%",background:"#ff4d4f",boxShadow:"0 0 5px #ff4d4f",flexShrink:0}}/>
                      <div style={{minWidth:0}}>
                        <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:fg}}>{r.title}</div>
                        <div style={{fontSize:10,color:text3}}>{r.ref} · {r.cat}</div>
                      </div>
                    </div>
                    <div style={{display:"flex",gap:5,flexShrink:0}}>
                      <button onClick={()=>setAiItem(r)} style={{background:"rgba(0,229,160,0.1)",border:"1px solid rgba(0,229,160,0.15)",color:"#00e5a0",padding:"3px 9px",borderRadius:5,cursor:"pointer",fontSize:10}}>AI</button>
                      <button onClick={()=>toggle(r.id,"compliant",r.title)} style={{background:"rgba(0,229,160,0.07)",border:"1px solid rgba(0,229,160,0.1)",color:"rgba(0,229,160,0.6)",padding:"3px 9px",borderRadius:5,cursor:"pointer",fontSize:10}}>✓</button>
                    </div>
                  </div>
                ))}
                {reqs.filter(r=>r.risk==="HIGH"&&statuses[r.id]!=="compliant").length===0&&(
                  <div style={{textAlign:"center",padding:"16px",color:"#00e5a0",fontSize:13}}>{t.noHighRisk}</div>
                )}
              </div>
            </div>
          )}

          {/* ───── CHECKLIST ───── */}
          {activeTab==="checklist"&&(
            <div>
              {/* Search + filters */}
              <div style={{display:"flex",gap:7,marginBottom:12,flexWrap:"wrap"}}>
                <input ref={searchRef} value={search} onChange={e=>setSearch(e.target.value)} placeholder={`${t.searchPlaceholder} (Press / to focus)`} style={{flex:1,minWidth:150,background:surface,border:`1px solid ${border}`,borderRadius:8,color:fg,padding:"8px 12px",fontSize:12,outline:"none"}}/>
                {["All","HIGH","MEDIUM","LOW"].map(r=>(
                  <button key={r} onClick={()=>setRiskFilter(r)} style={{padding:"7px 11px",borderRadius:7,border:"1px solid",borderColor:riskFilter===r?"rgba(255,255,255,0.25)":border,background:riskFilter===r?surface:"transparent",color:riskFilter===r?rc(r):text3,fontSize:10,cursor:"pointer",fontFamily:"'DM Mono',monospace",letterSpacing:.5}}>{r==="All"?"ALL":r}</button>
                ))}
                {/* Status filter */}
                <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)} style={{background:surface,border:`1px solid ${border}`,borderRadius:7,color:fg,padding:"7px 10px",fontSize:11,outline:"none",cursor:"pointer"}}>
                  {[["All","All Status"],["compliant",t.compliant],["in-progress",t.inProgress],["non-compliant",t.nonCompliant],["not-assessed",t.notAssessed]].map(([v,l])=>(
                    <option key={v} value={v} style={{background:"#0f1923"}}>{l}</option>
                  ))}
                </select>
                {/* Sort */}
                <select value={sortBy} onChange={e=>setSortBy(e.target.value)} style={{background:surface,border:`1px solid ${border}`,borderRadius:7,color:fg,padding:"7px 10px",fontSize:11,outline:"none",cursor:"pointer"}}>
                  <option value="risk" style={{background:"#0f1923"}}>{t.sortRisk}</option>
                  <option value="status" style={{background:"#0f1923"}}>{t.sortStatus}</option>
                  <option value="deadline" style={{background:"#0f1923"}}>{t.sortDeadline}</option>
                  <option value="alpha" style={{background:"#0f1923"}}>{t.sortAlpha}</option>
                </select>
              </div>
              {/* Category tabs */}
              <div style={{display:"flex",gap:5,marginBottom:12,flexWrap:"wrap"}}>
                {cats.map(cat=>(
                  <button key={cat} onClick={()=>setCatFilter(cat)} style={{padding:"4px 11px",borderRadius:14,border:"1px solid",borderColor:catFilter===cat?`${REG_COLOR[activeReg]}50`:border,background:catFilter===cat?`${REG_COLOR[activeReg]}12`:"transparent",color:catFilter===cat?REG_COLOR[activeReg]:text3,fontSize:11,cursor:"pointer",whiteSpace:"nowrap",transition:"all 0.15s"}}>{cat}</button>
                ))}
              </div>
              <div style={{fontSize:11,color:text3,marginBottom:9}}>{filtered.length}/{total} requirements</div>

              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {filtered.map(req=>{
                  const st=statuses[req.id],isEx=expanded===req.id,ev=evidence[req.id]||[];
                  return(
                    <div key={req.id} style={{background:surface,border:"1px solid",borderColor:st==="compliant"?"rgba(0,229,160,0.2)":st==="non-compliant"?"rgba(255,77,79,0.2)":st==="in-progress"?"rgba(250,173,20,0.18)":border,borderRadius:10,overflow:"hidden",transition:"border-color 0.2s"}} title={`${req.title} — ${req.ref}`}>
                      <div style={{padding:"11px 13px"}} className="rh">
                        <div style={{display:"flex",alignItems:"flex-start",gap:9}}>
                          {/* Status buttons */}
                          <div style={{display:"flex",gap:3,flexShrink:0,marginTop:1}}>
                            {[["compliant","✓","#00e5a0"],["in-progress","↻","#faad14"],["non-compliant","✗","#ff4d4f"]].map(([v,icon,col])=>(
                              <button key={v} onClick={()=>toggle(req.id,v,req.title)} title={v} style={{width:21,height:21,borderRadius:5,border:"1px solid",borderColor:st===v?col:isDark?"rgba(255,255,255,0.1)":"rgba(0,0,0,0.1)",background:st===v?col:"transparent",color:st===v?(v==="in-progress"?"#000":"#fff"):text3,cursor:"pointer",fontSize:9,fontWeight:900,display:"flex",alignItems:"center",justifyContent:"center",transition:"all 0.15s"}}>{icon}</button>
                            ))}
                          </div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap",marginBottom:3}}>
                              <span style={{fontSize:12,fontWeight:700,fontFamily:"'Syne',sans-serif",color:fg}}>{req.title}</span>
                              <span style={{fontSize:9,padding:"1px 5px",borderRadius:3,background:rb(req.risk),color:rc(req.risk),fontFamily:"'DM Mono',monospace"}}>{req.risk}</span>
                              <span style={{fontSize:9,padding:"1px 5px",borderRadius:3,color:text3,border:`1px solid ${border}`,fontFamily:"'DM Mono',monospace"}}>{req.ref}</span>
                              {deadlines[req.id]&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:3,background:"rgba(0,180,216,0.1)",color:"#00b4d8"}}>📅 {fmtDate(deadlines[req.id])}</span>}
                              {ev.length>0&&<span style={{fontSize:9,padding:"1px 5px",borderRadius:3,background:"rgba(168,85,247,0.1)",color:"#a855f7"}}>📎 {ev.length}</span>}
                            </div>
                            <p style={{fontSize:11,color:text2,lineHeight:1.5}}>{req.desc}</p>
                          </div>
                          <div style={{display:"flex",gap:4,flexShrink:0}}>
                            <button onClick={()=>setAiItem(req)} style={{background:"rgba(0,229,160,0.08)",border:"1px solid rgba(0,229,160,0.12)",color:"#00e5a0",padding:"4px 8px",borderRadius:6,cursor:"pointer",fontSize:10,whiteSpace:"nowrap"}}>{t.aiAdvisor}</button>
                            <button onClick={()=>setExpanded(isEx?null:req.id)} style={{background:surface,border:`1px solid ${border}`,color:text3,padding:"4px 7px",borderRadius:6,cursor:"pointer",fontSize:10}}>{isEx?"▲":"▼"}</button>
                          </div>
                        </div>
                      </div>
                      {isEx&&(
                        <div style={{padding:"0 13px 13px",borderTop:`1px solid ${border}`}}>
                          <div style={{background:"rgba(255,165,0,0.05)",borderRadius:7,padding:"7px 11px",marginBottom:10,marginTop:10}}>
                            <span style={{fontSize:10,color:"#faad14",fontWeight:600}}>→ {t.fix}: </span>
                            <span style={{fontSize:11,color:text2}}>{req.fix}</span>
                          </div>
                          <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:9,marginBottom:9}}>
                            <div>
                              <div style={{fontSize:9,color:text3,marginBottom:4,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{t.deadline.toUpperCase()}</div>
                              <input type="date" value={deadlines[req.id]||""} onChange={e=>saveDeadline(req.id,e.target.value,req.title)} style={{width:"100%",background:surface,border:`1px solid ${border}`,borderRadius:6,color:fg,padding:"6px 9px",fontSize:11,outline:"none"}}/>
                            </div>
                            <div>
                              <div style={{fontSize:9,color:text3,marginBottom:4,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{t.notes.toUpperCase()}</div>
                              {editNote===req.id?(
                                <div style={{display:"flex",gap:4}}>
                                  <input value={tempNote} onChange={e=>setTempNote(e.target.value)} onKeyDown={e=>e.key==="Enter"&&saveNote(req.id,req.title)} style={{flex:1,background:surface,border:`1px solid rgba(0,229,160,0.25)`,borderRadius:6,color:fg,padding:"6px 8px",fontSize:11,outline:"none"}} autoFocus placeholder="Add note..."/>
                                  <button onClick={()=>saveNote(req.id,req.title)} style={{background:"#00e5a0",border:"none",color:"#000",padding:"6px 9px",borderRadius:6,cursor:"pointer",fontSize:10,fontWeight:700}}>{t.saveNote}</button>
                                  <button onClick={()=>setEditNote(null)} style={{background:surface,border:`1px solid ${border}`,color:text3,padding:"6px 7px",borderRadius:6,cursor:"pointer",fontSize:10}}>✗</button>
                                </div>
                              ):(
                                <div onClick={()=>{setEditNote(req.id);setTempNote(notes[req.id]||"");}} style={{background:surface,border:`1px solid ${border}`,borderRadius:6,padding:"6px 9px",fontSize:11,color:notes[req.id]?text2:text3,cursor:"pointer",minHeight:30,lineHeight:1.5}}>
                                  {notes[req.id]||"Click to add notes..."}
                                </div>
                              )}
                            </div>
                          </div>
                          {/* Evidence */}
                          <div>
                            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                              <div style={{fontSize:9,color:text3,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>EVIDENCE ({ev.length})</div>
                              <button onClick={()=>{setNewEvidence(req.id);setTempEvidence("");}} style={{background:"rgba(0,180,216,0.1)",border:"1px solid rgba(0,180,216,0.2)",color:"#00b4d8",padding:"2px 8px",borderRadius:4,cursor:"pointer",fontSize:9}}>{t.addEvidence}</button>
                            </div>
                            {newEvidence===req.id&&(
                              <div style={{display:"flex",gap:4,marginBottom:6}}>
                                <input value={tempEvidence} onChange={e=>setTempEvidence(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addEvidenceItem(req.id,req.title)} style={{flex:1,background:surface,border:"1px solid rgba(0,180,216,0.25)",borderRadius:6,color:fg,padding:"5px 8px",fontSize:11,outline:"none"}} autoFocus placeholder="Evidence reference, document name, URL..."/>
                                <button onClick={()=>addEvidenceItem(req.id,req.title)} style={{background:"#00b4d8",border:"none",color:"#000",padding:"5px 9px",borderRadius:6,cursor:"pointer",fontSize:10,fontWeight:700}}>✓</button>
                                <button onClick={()=>setNewEvidence(null)} style={{background:surface,border:`1px solid ${border}`,color:text3,padding:"5px 7px",borderRadius:6,cursor:"pointer",fontSize:10}}>✗</button>
                              </div>
                            )}
                            {ev.map((e2,i)=>(
                              <div key={i} style={{display:"flex",gap:7,alignItems:"flex-start",padding:"4px 0",borderBottom:i<ev.length-1?`1px solid ${border}`:"none"}}>
                                <span style={{color:"#00b4d8",fontSize:10,flexShrink:0}}>📎</span>
                                <div style={{flex:1}}>
                                  <div style={{fontSize:11,color:text2}}>{e2.text}</div>
                                  <div style={{fontSize:9,color:text3,marginTop:1}}>{fmtTs(e2.ts)}</div>
                                </div>
                              </div>
                            ))}
                          </div>
                          {/* mailto reminder */}
                          {deadlines[req.id]&&(
                            <div style={{marginTop:9}}>
                              <a href={`mailto:?subject=Compliance Deadline: ${req.title}&body=Reminder: "${req.title}" (${req.ref}) deadline is ${fmtDate(deadlines[req.id])}.\n\nRequired action: ${req.fix}\n\n-- Sent from ComplianceOS`} style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(0,180,216,0.08)",border:"1px solid rgba(0,180,216,0.15)",color:"#00b4d8",padding:"5px 11px",borderRadius:6,fontSize:10,textDecoration:"none"}}>
                                ✉️ {t.remindTeam}
                              </a>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ───── TEAM ───── */}
          {activeTab==="team"&&(
            <div>
              <div style={{marginBottom:18}}>
                <h2 style={{fontSize:17,fontFamily:"'Syne',sans-serif",fontWeight:900,marginBottom:4,color:fg}}>{t.team} Workspace</h2>
                <p style={{color:text3,fontSize:12}}>Shared activity feed visible to all team members</p>
              </div>
              <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:14,marginBottom:14}}>
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"18px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>{t.yourProfile}</div>
                  {[["YOUR NAME",userName,setUserName,"username"],["COMPANY",companyName,setCompanyName,"company"]].map(([label,val,set,key])=>(
                    <div key={key} style={{marginBottom:10}}>
                      <label style={{fontSize:9,color:text3,display:"block",marginBottom:4,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{label}</label>
                      <input value={val} onChange={e=>set(e.target.value)} onBlur={()=>sSet(key,val)} style={{width:"100%",background:isDark?"rgba(255,255,255,0.05)":"#f7fafc",border:`1px solid ${border}`,borderRadius:7,color:fg,padding:"7px 10px",fontSize:12,outline:"none"}}/>
                    </div>
                  ))}
                </div>
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"18px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>{t.postUpdate}</div>
                  <textarea value={teamMsg} onChange={e=>setTeamMsg(e.target.value)} placeholder="Post update for team..." style={{width:"100%",background:isDark?"rgba(255,255,255,0.04)":"#f7fafc",border:`1px solid ${border}`,borderRadius:7,color:fg,padding:"8px 10px",fontSize:12,outline:"none",resize:"none",height:70,fontFamily:"'Space Grotesk',sans-serif"}}/>
                  <button onClick={async()=>{if(teamMsg.trim()){await addTeamFeed(teamMsg.trim());setTeamMsg("");}}} disabled={!teamMsg.trim()} style={{marginTop:7,width:"100%",padding:"8px",borderRadius:7,border:"none",background:teamMsg.trim()?"rgba(0,229,160,0.15)":surface,color:teamMsg.trim()?"#00e5a0":text3,cursor:teamMsg.trim()?"pointer":"default",fontSize:12,fontWeight:700}}>
                    {t.postUpdate}
                  </button>
                </div>
              </div>
              <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"18px"}}>
                <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>{t.teamActivity}</div>
                {teamFeed.length===0?(<div style={{textAlign:"center",padding:"24px",color:text3,fontSize:12}}>No activity yet. Complete requirements or post updates.</div>)
                :teamFeed.map((f,i)=>(
                  <div key={f.id} style={{display:"flex",gap:10,alignItems:"flex-start",padding:"8px 0",borderBottom:i<teamFeed.length-1?`1px solid ${border}`:"none"}}>
                    <div style={{width:26,height:26,borderRadius:7,background:"linear-gradient(135deg,#00e5a0,#00b4d8)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:900,color:"#000",flexShrink:0}}>{f.user?.charAt(0)||"?"}</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:12,color:text2,lineHeight:1.5}}>{f.msg}</div>
                      <div style={{fontSize:10,color:text3,marginTop:2}}>{f.user} · {f.company} · {fmtTs(f.ts)}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* ───── AUDIT ───── */}
          {activeTab==="audit"&&(
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18,flexWrap:"wrap",gap:8}}>
                <div>
                  <h2 style={{fontSize:17,fontFamily:"'Syne',sans-serif",fontWeight:900,marginBottom:4,color:fg}}>{t.audit}</h2>
                  <p style={{color:text3,fontSize:12}}>Complete log of all compliance actions — {auditLog.length} entries</p>
                </div>
                <button onClick={()=>{const r=auditLog.map(a=>({Timestamp:fmtTs(a.ts),User:a.user||"",Action:a.action,Requirement:a.item,Detail:a.detail||""}));const ws=XLSX.utils.json_to_sheet(r);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Audit");XLSX.writeFile(wb,"audit-log.xlsx");}} style={{background:"rgba(0,229,160,0.1)",border:"1px solid rgba(0,229,160,0.2)",color:"#00e5a0",padding:"7px 14px",borderRadius:8,cursor:"pointer",fontSize:12}}>Export Log</button>
              </div>
              {auditLog.length===0?(<div style={{textAlign:"center",padding:"40px",color:text3,fontSize:13}}><div style={{fontSize:28,marginBottom:10}}>⊙</div>Actions will appear here automatically.</div>)
              :auditLog.map((e2,i)=>(
                <div key={e2.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",background:surface,border:`1px solid ${border}`,borderRadius:8,marginBottom:4}}>
                  <div style={{width:7,height:7,borderRadius:"50%",flexShrink:0,background:e2.action.includes("Compliant")&&!e2.action.includes("Non")?"#00e5a0":e2.action.includes("Non")?"#ff4d4f":e2.action.includes("Progress")?"#faad14":"#00b4d8"}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <span style={{fontSize:12,fontWeight:600,color:e2.action.includes("Compliant")&&!e2.action.includes("Non")?"#00e5a0":e2.action.includes("Non")?"#ff4d4f":e2.action.includes("Progress")?"#faad14":fg}}>{e2.action}</span>
                    <span style={{fontSize:12,color:text2,marginLeft:7}}>— {e2.item}</span>
                    {e2.detail&&<span style={{fontSize:10,color:text3,marginLeft:5}}>({e2.detail})</span>}
                    {e2.user&&<div style={{fontSize:9,color:text3,marginTop:1}}>by {e2.user}</div>}
                  </div>
                  <span style={{fontSize:10,color:text3,fontFamily:"'DM Mono',monospace",flexShrink:0,whiteSpace:"nowrap"}}>{fmtTs(e2.ts)}</span>
                </div>
              ))}
            </div>
          )}

          {/* ───── REPORTS ───── */}
          {activeTab==="reports"&&(
            <div>
              <div style={{marginBottom:18}}>
                <h2 style={{fontSize:17,fontFamily:"'Syne',sans-serif",fontWeight:900,marginBottom:4,color:fg}}>{t.reports}</h2>
                <p style={{color:text3,fontSize:12}}>Export for auditors, legal teams, regulators</p>
              </div>
              <div style={{background:`linear-gradient(135deg,rgba(0,229,160,0.05),rgba(0,180,216,0.03))`,border:"1px solid rgba(0,229,160,0.1)",borderRadius:16,padding:"22px 24px",marginBottom:14}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18,flexWrap:"wrap",gap:10}}>
                  <div>
                    <div style={{fontSize:10,color:text3,letterSpacing:1.5,fontFamily:"'DM Mono',monospace",marginBottom:6}}>{activeReg} REPORT</div>
                    <div style={{fontSize:18,fontWeight:900,fontFamily:"'Syne',sans-serif",color:fg}}>{companyName}</div>
                    <div style={{fontSize:11,color:text3,marginTop:3}}>{new Date().toLocaleDateString("en-GB",{day:"numeric",month:"long",year:"numeric"})} · {userName}</div>
                    <div style={{fontSize:10,color:"rgba(255,77,79,0.8)",marginTop:5,display:"inline-block",background:"rgba(255,77,79,0.07)",padding:"2px 8px",borderRadius:4}}>⚠️ Guidance only</div>
                  </div>
                  <ScoreGauge score={score} size={96}/>
                </div>
                <div style={{display:"grid",gridTemplateColumns:mobile?"1fr 1fr":"repeat(4,1fr)",gap:9}}>
                  {[{l:t.compliant,v:compliant,c:"#00e5a0"},{l:t.inProgress,v:inProgress,c:"#faad14"},{l:t.nonCompliant,v:nonCompliant,c:"#ff4d4f"},{l:t.notAssessed,v:total-compliant-inProgress-nonCompliant,c:text3}].map((s,i)=>(
                    <div key={i} style={{background:"rgba(255,255,255,0.04)",borderRadius:8,padding:"11px",textAlign:"center"}}>
                      <div style={{fontSize:22,fontWeight:900,fontFamily:"'Syne',sans-serif",color:s.c}}>{s.v}</div>
                      <div style={{fontSize:10,color:text3,marginTop:3}}>{s.l}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"repeat(4,1fr)",gap:11,marginBottom:14}}>
                {[
                  {label:t.exportPDF,icon:"📄",color:"#00e5a0",bg:"rgba(0,229,160,0.06)",border:"rgba(0,229,160,0.2)",action:exportPDF,sub:"Opens print dialog"},
                  {label:t.exportExcel,icon:"📊",color:"#00b4d8",bg:"rgba(0,180,216,0.06)",border:"rgba(0,180,216,0.2)",action:exportExcel,sub:"Full data + audit log"},
                  {label:t.exportCSV,icon:"📋",color:"#f59e0b",bg:"rgba(245,158,11,0.06)",border:"rgba(245,158,11,0.2)",action:exportCSV,sub:"Import into any tool"},
                  {label:t.shareReport,icon:"🔗",color:"#a855f7",bg:"rgba(168,85,247,0.06)",border:"rgba(168,85,247,0.2)",action:shareReport,sub:shareMsg||"Copy to clipboard"},
                ].map((btn,i)=>(
                  <button key={i} onClick={btn.action} style={{padding:"16px",borderRadius:11,border:`1px solid ${btn.border}`,background:btn.bg,color:btn.color,cursor:"pointer",fontSize:13,fontWeight:800,fontFamily:"'Syne',sans-serif",textAlign:"center"}}>
                    <div style={{fontSize:20,marginBottom:5}}>{btn.icon}</div>
                    {btn.label}
                    <div style={{fontSize:10,fontWeight:400,opacity:0.6,marginTop:2}}>{btn.sub}</div>
                  </button>
                ))}
              </div>
              {/* CSV Import */}
              <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:10,color:fg}}>{t.importCSV}</div>
                <div style={{fontSize:12,color:text3,marginBottom:10}}>Import status updates from a CSV file (exported from ComplianceOS).</div>
                <label style={{display:"inline-flex",alignItems:"center",gap:8,padding:"8px 16px",borderRadius:8,border:`1px solid ${border}`,background:surface,color:text2,cursor:"pointer",fontSize:12}}>
                  📂 Choose CSV File
                  <input type="file" accept=".csv" onChange={e=>importCSV(e.target.files?.[0])} style={{display:"none"}}/>
                </label>
                {csvImportMsg&&<span style={{marginLeft:12,fontSize:12,color:"#00e5a0"}}>{csvImportMsg}</span>}
              </div>
            </div>
          )}

          {/* ───── SETTINGS ───── */}
          {activeTab==="settings"&&(
            <div>
              <h2 style={{fontSize:17,fontFamily:"'Syne',sans-serif",fontWeight:900,marginBottom:16,color:fg}}>{t.settings}</h2>
              <div style={{display:"flex",flexDirection:"column",gap:12}}>

                {/* Theme + Language */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>Appearance</div>
                  <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:10}}>
                    <div>
                      <label style={{fontSize:9,color:text3,display:"block",marginBottom:5,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{t.theme.toUpperCase()}</label>
                      <button onClick={toggleTheme} style={{width:"100%",padding:"9px",borderRadius:8,border:`1px solid ${border}`,background:surface,color:fg,cursor:"pointer",fontSize:12,fontWeight:600}}>
                        {isDark?"☀️ "+t.lightMode:"🌙 "+t.darkMode}
                      </button>
                    </div>
                    <div>
                      <label style={{fontSize:9,color:text3,display:"block",marginBottom:5,fontFamily:"'DM Mono',monospace",letterSpacing:1}}>{t.language.toUpperCase()}</label>
                      <select value={lang} onChange={e=>{setLang(e.target.value);sSet("lang",e.target.value);}} style={{width:"100%",background:isDark?"rgba(255,255,255,0.05)":"#f7fafc",border:`1px solid ${border}`,borderRadius:8,color:fg,padding:"9px 10px",fontSize:12,outline:"none",cursor:"pointer"}}>
                        <option value="en">🇬🇧 English</option>
                        <option value="de">🇩🇪 Deutsch</option>
                        <option value="fr">🇫🇷 Français</option>
                      </select>
                    </div>
                  </div>
                </div>

                {/* Notifications */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:6,color:fg}}>{t.notifications}</div>
                  <p style={{fontSize:12,color:text3,marginBottom:10}}>Get browser alerts 7 days before compliance deadlines.</p>
                  <button onClick={enableNotifications} style={{padding:"8px 18px",borderRadius:8,border:`1px solid ${notifEnabled?"rgba(0,229,160,0.3)":border}`,background:notifEnabled?"rgba(0,229,160,0.1)":surface,color:notifEnabled?"#00e5a0":text2,cursor:"pointer",fontSize:12,fontWeight:600}}>
                    {notifEnabled?"✓ Alerts Enabled":t.enableNotif}
                  </button>
                </div>

                {/* Regulation updates */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:4,color:fg}}>🔔 Regulation Update Monitor</div>
                  <p style={{fontSize:12,color:text3,marginBottom:10}}>AI-powered check for recent regulatory changes.</p>
                  <button onClick={checkRegUpdates} disabled={checkingUpdates} style={{padding:"8px 16px",borderRadius:8,border:"1px solid rgba(0,229,160,0.2)",background:"rgba(0,229,160,0.08)",color:"#00e5a0",cursor:checkingUpdates?"default":"pointer",fontSize:12,fontWeight:700,marginBottom:regUpdateResult?10:0}}>
                    {checkingUpdates?"Checking...":"Check for Updates"}
                  </button>
                  {checkingUpdates&&<div style={{display:"flex",alignItems:"center",gap:6,marginTop:8,color:text3,fontSize:11}}><div style={{display:"flex",gap:4}}>{[0,1,2].map(i=><div key={i} style={{width:5,height:5,borderRadius:"50%",background:"#00e5a0",animation:`dot 1.2s ease-in-out ${i*0.2}s infinite`}}/>)}</div>Analysing...</div>}
                  {regUpdateResult&&<pre style={{fontSize:11,color:text2,lineHeight:1.7,whiteSpace:"pre-wrap",background:isDark?"rgba(255,255,255,0.02)":"#f7fafc",borderRadius:8,padding:"10px 12px",fontFamily:"'Space Grotesk',sans-serif",marginTop:10}}>{regUpdateResult}</pre>}
                </div>

                {/* Integrations */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>{t.integrations}</div>
                  <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:8}}>
                    {[{n:"Slack",d:"Compliance alerts in channels",i:"💬"},{n:"Jira",d:"Auto-create tickets for issues",i:"🎯"},{n:"Microsoft Teams",d:"Team notifications",i:"👥"},{n:"Zapier",d:"Connect 5000+ apps",i:"⚡"}].map(int=>(
                      <div key={int.n} style={{background:isDark?"rgba(255,255,255,0.03)":"#f7fafc",border:`1px solid ${border}`,borderRadius:9,padding:"10px 13px",display:"flex",alignItems:"center",gap:9}}>
                        <span style={{fontSize:18}}>{int.i}</span>
                        <div style={{flex:1}}>
                          <div style={{fontSize:12,fontWeight:600,color:fg}}>{int.n}</div>
                          <div style={{fontSize:10,color:text3}}>{int.d}</div>
                        </div>
                        <div style={{fontSize:9,padding:"2px 7px",borderRadius:3,background:"rgba(250,173,20,0.1)",color:"#faad14",fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap"}}>{t.comingSoon}</div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Keyboard shortcuts */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:10,color:fg}}>{t.keyboardShortcuts}</div>
                  <div style={{display:"grid",gridTemplateColumns:mobile?"1fr":"1fr 1fr",gap:6}}>
                    {[t.shortcutSearch,t.shortcutUndo,t.shortcutDark,t.shortcutAudit].map(s=>(
                      <div key={s} style={{background:isDark?"rgba(255,255,255,0.03)":"#f7fafc",borderRadius:7,padding:"7px 11px",fontSize:11,color:text2,fontFamily:"'DM Mono',monospace"}}>{s}</div>
                    ))}
                  </div>
                </div>

                {/* Data management */}
                <div style={{background:surface,border:`1px solid ${border}`,borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:12,color:fg}}>{t.dataManagement}</div>
                  <div style={{display:"flex",gap:9,flexWrap:"wrap"}}>
                    <button onClick={()=>{const d={statuses,deadlines,notes,evidence,auditLog,companyName,userName,version:"4.0",exported:new Date().toISOString()};const b=new Blob([JSON.stringify(d,null,2)],{type:"application/json"});const u=URL.createObjectURL(b);const a=document.createElement("a");a.href=u;a.download=`cos-backup-${new Date().toISOString().split("T")[0]}.json`;a.click();}} style={{padding:"8px 15px",borderRadius:8,border:"1px solid rgba(0,229,160,0.2)",background:"rgba(0,229,160,0.07)",color:"#00e5a0",cursor:"pointer",fontSize:12,fontWeight:700}}>↓ {t.exportBackup}</button>
                    <button onClick={async()=>{if(confirm("Reset ALL data? Cannot be undone.")){dispatch({type:"SET",payload:{}});setDeadlines({});setNotes({});setEvidence({});setAuditLog([]);await Promise.all([sSet("statuses",{}),sSet("deadlines",{}),sSet("notes",{}),sSet("evidence",{}),sSet("audit",[])]);setShowTour(true);}}} style={{padding:"8px 15px",borderRadius:8,border:"1px solid rgba(255,77,79,0.2)",background:"rgba(255,77,79,0.06)",color:"#ff4d4f",cursor:"pointer",fontSize:12,fontWeight:700}}>⚠ {t.resetData}</button>
                    <button onClick={()=>setShowTour(true)} style={{padding:"8px 15px",borderRadius:8,border:`1px solid ${border}`,background:surface,color:text2,cursor:"pointer",fontSize:12}}>▶ Replay Tour</button>
                  </div>
                  <div style={{fontSize:11,color:text3,marginTop:10,lineHeight:1.7}}>
                    Audit entries: <strong style={{color:text2}}>{auditLog.length}</strong> · Total requirements: <strong style={{color:text2}}>{Object.values(REGS).flat().length}</strong> · Storage: Cloud persistent
                  </div>
                </div>

                {/* Legal */}
                <div style={{background:"rgba(255,77,79,0.04)",border:"1px solid rgba(255,77,79,0.12)",borderRadius:13,padding:"16px 20px"}}>
                  <div style={{fontSize:13,fontWeight:700,fontFamily:"'Syne',sans-serif",marginBottom:9,color:"#ff6b6b"}}>⚠️ Legal Notices</div>
                  <div style={{fontSize:12,color:text2,lineHeight:1.8}}>
                    <p>• Compliance guidance only — <strong style={{color:fg}}>NOT legal advice</strong></p>
                    <p>• Scores are indicative — <strong style={{color:fg}}>do not guarantee compliance</strong></p>
                    <p>• Always consult a <strong style={{color:fg}}>qualified legal professional</strong></p>
                    <p>• Verify against: <strong style={{color:fg}}>EUR-Lex · ICO · FTC · HHS.gov</strong></p>
                    <p>• ComplianceOS accepts <strong style={{color:fg}}>no liability</strong> for fines or penalties</p>
                  </div>
                  <div style={{display:"flex",gap:8,marginTop:10}}>
                    <button onClick={()=>setShowPrivacy(true)} style={{padding:"6px 13px",borderRadius:7,border:`1px solid ${border}`,background:surface,color:text3,cursor:"pointer",fontSize:11}}>{t.privacyPolicy}</button>
                    <button onClick={()=>setShowTerms(true)} style={{padding:"6px 13px",borderRadius:7,border:`1px solid ${border}`,background:surface,color:text3,cursor:"pointer",fontSize:11}}>{t.termsOfService}</button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Footer */}
          <div style={{marginTop:24,paddingTop:14,borderTop:`1px solid ${border}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:7}}>
            <span style={{fontSize:10,color:text3}}>⚠️ {t.disclaimer}</span>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setShowPrivacy(true)} style={{background:"none",border:"none",color:text3,fontSize:10,cursor:"pointer"}}>{t.privacyPolicy}</button>
              <button onClick={()=>setShowTerms(true)} style={{background:"none",border:"none",color:text3,fontSize:10,cursor:"pointer"}}>{t.termsOfService}</button>
            </div>
          </div>
        </div>
      </div>

      {/* Overlays */}
      {showTour&&<OnboardingTour lang={lang} onFinish={finishTour}/>}
      {!cookieOk&&<CookieBanner t={t} onAccept={acceptCookie}/>}
      {aiItem&&<AIAdvisor item={aiItem} reg={activeReg} lang={lang} onClose={()=>setAiItem(null)}/>}
      {showPolicy&&<PolicyGenerator company={companyName} lang={lang} onClose={()=>setShowPolicy(false)}/>}
      {showPrivacy&&<LegalModal title={t.privacyPolicy} content={`PRIVACY POLICY — ComplianceOS v4.0\n\n1. All compliance data stored in persistent cloud storage — never on our servers.\n2. AI Advisor sends only public regulatory requirement text to Anthropic API.\n3. No personal data transmitted. No tracking. No advertising cookies.\n4. You control all your data. Delete anytime via Settings > Reset All Data.\n5. Contact your organisation's DPO for privacy queries.\n\n⚠️ Template guidance only — review with legal counsel before use.`} onClose={()=>setShowPrivacy(false)}/>}
      {showTerms&&<LegalModal title={t.termsOfService} content={`TERMS OF SERVICE — ComplianceOS v4.0\n\n⚠️ CRITICAL: This tool provides compliance GUIDANCE ONLY — NOT legal advice.\n\n1. NOT LEGAL ADVICE: Always consult a qualified legal professional.\n2. LIMITATION OF LIABILITY: ComplianceOS accepts NO liability for fines, penalties, or enforcement actions.\n3. ACCURACY: Scores are indicative only. Verify against official sources.\n4. YOUR RESPONSIBILITY: You are solely responsible for actual regulatory compliance.\n5. AI CONTENT: Policy Generator output must be reviewed by a lawyer before use.\n6. GOVERNING LAW: Laws of England and Wales.\n\nBy using ComplianceOS you accept these terms.`} onClose={()=>setShowTerms(false)}/>}
    </div>
  );
}
