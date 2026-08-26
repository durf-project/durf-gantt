"""Canonical DURF roadmap content, transcribed from the DURF presentation (slides 5-10).

This module is the seed for `data/durf-roadmap.xlsx`. After the workbook exists it
becomes the source of truth -- edit the workbook, not this file. Re-running
`seed_workbook.py` overwrites the workbook, so only do that to start over.
"""

# SURF PowerPoint theme accents (ppt/theme/theme1.xml, clrScheme "SURF").
# The deck reuses accent1/accent2 for themes 5 and 6; a Gantt needs six
# distinguishable colours, so 5 and 6 take the remaining SURF accents.
SETTINGS = [
    ("project_title", "DURF", "Shown as the main heading."),
    ("project_subtitle", "Dutch Repository Federation — project roadmap",
     "Shown under the main heading."),
    ("start_year_month", "2026-01",
     "Calendar month that project month 1 maps to. Format YYYY-MM."),
    ("total_months", 48, "Project length in months. Drives the timeline width."),
    ("intro",
     "DURF reinforces the Netherlands Research Portal and the institutional "
     "repository/CRIS landscape so Dutch research output is more robust, "
     "sovereign and findable. Six themes run in parallel over four years.",
     "Short paragraph under the title. Leave blank to hide."),
    ("source_note",
     "Source: DURF presentation, 19 March 2026 — theme slides 5–10.",
     "Footer provenance line."),
]

THEMES = [
    dict(
        no=1,
        name="Governance of the Federation and Netherlands Research Portal",
        short="Governance",
        color="#EE7628",
        goal="A sustainable governance framework established for the Dutch federated "
             "research information ecosystem including parties maintaining "
             "Repository/CRIS systems, eDepot and OpenAIRE Graph.",
        lead="TU Delft",
        partners="SURF (community management, legal advice, project assistance)",
    ),
    dict(
        no=2,
        name="Metadata Quality Improvement",
        short="Metadata Quality",
        color="#F6DA44",
        goal="Repository/CRIS systems have adopted the updated metadata- and "
             "exchange-standards for high-quality research metadata that comply with "
             "international guidelines and meet national requirements.",
        lead="Leiden",
        partners="WUR (tech lead), SURF (community management, legal advice, project "
                 "assistance, architect), KB (WISH)",
    ),
    dict(
        no=3,
        name="Full Text Capture",
        short="Full Text Capture",
        color="#177ABF",
        goal="Comprehensive full-text availability in Dutch repository/CRIS systems "
             "for enhancing primarily preservation, and secondary text-mining, "
             "AI-usage and Taverne-mechanisms.",
        lead="UvA",
        partners="WUR (tech lead), SURF (community management, legal advice, project "
                 "assistance)",
    ),
    dict(
        no=4,
        name="KB National Library e-Depot Archiving",
        short="e-Depot Archiving",
        color="#2CA055",
        goal="Comprehensive preservation achieved of Dutch scientific output in the "
             "KB e-Depot.",
        lead="KB National Library",
        partners="WUR (tech lead), SURF (legal)",
    ),
    dict(
        no=5,
        name="Metadata Distribution & Discovery Optimization",
        short="Distribution & Discovery",
        color="#7B2882",
        goal="Increase the visibility of Dutch research content in major "
             "international indexes.",
        lead="Leiden",
        partners="WUR (tech lead)",
    ),
    dict(
        no=6,
        name="NL Research Portal Update",
        short="NL Research Portal",
        color="#E7303A",
        goal="Deliver a fully-functional, user-centered Netherlands Research Portal "
             "that serves as the primary discovery point for Dutch research outputs.",
        lead="TU Delft",
        partners="",
    ),
]

GORC = ("GORC Commons Model template :: https://view.officeapps.live.com/op/view.aspx"
        "?src=https%3A%2F%2Fwww.rd-alliance.org%2Fwp-content%2Fuploads%2F2024%2F07%2F"
        "GORC-International-Model-WG-Commons-Model-V1.1.xlsx&wdOrigin=BROWSELINK")
DASH = ("CRIS/Repo Dashboard :: "
        "https://surf-ori.github.io/dashboards/cris-repository-overview.html")

# theme, activity, outcomes, tags, start, end, milestones, links, notes
ACTIVITIES = [
    # ---- Theme 1 ---------------------------------------------------------
    (1, "Develop and secure signatures for a position paper on the shared "
        "information chain",
     "A position paper on the importance of the Shared Information Chain is signed "
     "by stakeholders.",
     "report", 1, 12, "", GORC, ""),
    (1, "Establish NaMeCo (Dutch National Membership Consortium for OpenAIRE "
        "representation)",
     "Signed federation / consortium agreement.",
     "signed agreement", 1, 12, "",
     "NaMeCo template :: https://www.openaire.eu/public-documents/nameco-template", ""),
    (1, "Draft and finalize formal agreements on roles and responsibilities between "
        "SURF, NaMeCo institutions, KB National Library, publishers, and OpenAIRE",
     "Draft & Signed Federation Rule Book.",
     "signed agreement", 1, 18, "", GORC, ""),
    (1, "Define legal structure for CC0 metadata influx",
     "Legal Guideline “metadata re-use” in CRISes, used in negotiations for "
     "future publisher deals, Scopus & Pure.",
     "legal", 6, 6, "6", "", ""),
    (1, "Create a comprehensive governance framework for all project themes",
     "Project governance framework.",
     "report", 1, 12, "", "", ""),
    (1, "Organize annual National Symposium and General Assembly for federation "
        "members",
     "4 National symposia.",
     "event", 12, 48, "12, 24, 36, 48", "",
     "Budget: € 23.000 per symposium × 4."),
    (1, "Review and update governance framework to ensure post-project sustainability",
     "Updated Federation Rule Book.",
     "report", 36, 48, "", "", ""),

    # ---- Theme 2 ---------------------------------------------------------
    (2, "Form and formalize NL Research Information Content Board through "
        "EduStandaard",
     "1 WISH 2.0 formalized at Edustandaard.nl.",
     "signed agreement", 1, 3, "",
     "EduStandaard werkgroepen :: https://www.edustandaard.nl/standaard_werkgroepen/",
     ""),
    (2, "Develop updated metadata application profile; translating current "
        "NL-DIDL-MODS requirements to OpenAIRE CERIF and DC guidelines",
     "1 oai_cerif_openaire_NL and 1 oai_dc_openaire_NL application profile published "
     "at Edustandaard.nl.",
     "standard", 3, 12, "",
     "NL-DIDL-MODS :: https://www.edustandaard.nl/standaard_afspraken/?edu_orderby=title"
     "&werkgebieden%5B%5D=ho&registratiestatussen%5B%5D=geregistreerd"
     "&registratiestatussen%5B%5D=aangemeld&gebruiksadviezen%5B%5D=verplicht"
     "&gebruiksadviezen%5B%5D=aangeraden&gebruiksadviezen%5B%5D=onder-voorwaarden"
     "&gebruiksadviezen%5B%5D=onbekend"
     " | OpenAIRE CERIF and DC guidelines :: https://guidelines.openaire.eu/"
     " | RSL: Really Simple Licensing :: https://rslstandard.org/", ""),
    (2, "Support repositories, CRIS systems, and KB eDepot in adopting the updated "
        "profile",
     "16 CRIS systems & 31 institutional repositories adopted the new application "
     "profile.",
     "configuration", 12, 18, "", "", ""),
    (2, "Monitor compliance with the new standard and share implementation best "
        "practices",
     "1 OAI Compliancy Validator. | 1 Compliancy health CRIS/Repo Dashboard: chart "
     "with validation results per repo over time. | 4 WISH 2.0 best practice "
     "workshops.",
     "software, service, training", 12, 48, "", DASH, ""),
    (2, "Develop proof of concept for enriching local metadata with the OpenAIRE data "
        "broker",
     "15 CRIS/repo managers attend training on the OpenAIRE broker feed to ingest "
     "enriched metadata in CRIS/repo. | Evaluation report on feedback quality.",
     "training, software, report", 12, 24, "",
     "OpenAIRE Broker :: https://catalogue.openaire.eu/service/openaire.broker/overview"
     " | Example broker feed :: "
     "https://provide.openaire.eu/repository/eurocrisdris::01222/events", ""),
    (2, "Monitor usage of core entity PIDs (ROR, ORCiD, DOI) and promote local "
        "implementation",
     "1 PID health CRIS/Repo Dashboard: chart with ROR, ORCiD and DOI presence in "
     "records per repo over time.",
     "software, service", 12, 48, "", DASH, ""),
    (2, "Launch download statistics sharing pilot on OpenAIRE",
     "2 CRIS/repos enabled for OpenAIRE Usage Counts.",
     "configuration", 12, 18, "",
     "OpenAIRE Usage Counts :: "
     "https://catalogue.openaire.eu/service/openaire.usagecounts/overview", ""),

    # ---- Theme 3 ---------------------------------------------------------
    (3, "Establish baseline metrics and reporting system for full-text uptake",
     "1 Digital sovereignty health CRIS/Repo Dashboard: chart with ratio of metadata "
     "records containing a link to a file (PDF) on the institutional domain, per repo "
     "over time.",
     "software, service", 1, 6, "", DASH, ""),
    (3, "Create inventory of existing tools and methods for full-text capture",
     "1 advisory report with an inventory of existing tools and methods for full-text "
     "capture, by desk research & interviews with WISH, the PURE user group and OA "
     "Switchboard.",
     "report", 1, 12, "", "", ""),
    (3, "Develop and pilot a local full-text capture service with shareable code and "
        "practices",
     "1 selected full-text capture method made available for other repos. | "
     "1 documentation set for others to implement.",
     "software, report", 12, 24, "", "", ""),
    (3, "Implement ecosystem-wide improvements to increase full-text availability",
     "2 workshops for CRIS/repo managers to implement the selected method.",
     "training", 24, 48, "", "", ""),

    # ---- Theme 4 ---------------------------------------------------------
    (4, "Implement pull mechanism for metadata and full-text publications, partly via "
        "OpenAIRE",
     "1 added usage term in provide.openaire; full-text download for archiving by the "
     "KB e-Depot. | 1 weekly metadata + full-text harvest package of 13 CRISes and 31 "
     "institutional repositories.",
     "legal, feature, data", 3, 12, "", "", ""),
    (4, "Archive and back up all Dutch CRIS and repository content at KB National "
        "Library e-Depot",
     "3.2 million records and full text safely stored in the KB e-Depot.",
     "service", 12, 48, "", "", ""),
    (4, "Establish monitoring system for preservation status of publications and "
        "metadata",
     "1 Backup health CRIS/Repo Dashboard: chart with ratio of PDFs in the e-Depot vs "
     "in the repository, per repo over time.",
     "software, service", 12, 48, "", DASH,
     "Slide timeline bar suggested months 3–18; the written months (12–48) "
     "are used here. Confirm with the theme lead."),
    (4, "Update and maintain URN:NBN resolver functionality",
     "URN:NBN resolver.",
     "service", 12, 18, "", "",
     "Slide timeline bar suggested months 1–6; the written months (12–18) "
     "are used here. Confirm with the theme lead."),

    # ---- Theme 5 ---------------------------------------------------------
    (5, "Develop mappings and best practices for major indexes to improve Dutch "
        "research visibility",
     "1 guideline for adding repository content to OpenAlex, WorldCat and Google "
     "Scholar.",
     "report", 1, 12, "",
     "Google Scholar inclusion guidelines :: "
     "https://scholar.google.com/intl/en/scholar/inclusion.html", ""),
    (5, "Implement DOI-minting capabilities for grey literature in DURF repositories",
     "13 CRISes + 31 repositories have minted a DOI for records that did not have a "
     "DOI but do have a PDF attached locally.",
     "configuration, software", 6, 18, "", "", ""),
    (5, "Create monitoring system for tracking DURF content visibility across major "
        "indexes",
     "1 Discoverability health CRIS/Repo Dashboard: chart indicating implementation of "
     "the Google Scholar index guidelines, per repo over time.",
     "software, service", 12, 48, "",
     DASH + " | Google Scholar inclusion guidelines :: "
     "https://scholar.google.com/intl/en/scholar/inclusion.html",
     "Slide timeline column also noted months 1–18; the written months "
     "(12–48) are used here. Confirm with the theme lead."),

    # ---- Theme 6 ---------------------------------------------------------
    (6, "Establish Netherlands Research Portal Steering Committee with rotating NaMeCo "
        "membership",
     "NL Portal steering committee.",
     "signed agreement", 1, 25, "1, 25", "",
     "Milestone-driven: convened in month 1 and refreshed in month 25."),
    (6, "Conduct regular stakeholder needs assessments",
     "Customer value report on Portal needs / business cases.",
     "report", 6, 42, "6, 18, 30, 42", "", "Repeats every 12 months."),
    (6, "Create development roadmaps aligned with OpenAIRE Connect and project themes",
     "Development roadmap.",
     "report", 9, 45, "9, 21, 33, 45", "", "Repeats every 12 months."),
    (6, "Release quarterly portal updates with major versions annually",
     "4 releases of the portal.",
     "software, service", 12, 48, "12, 24, 36, 48", "",
     "Quarterly updates; the milestones mark the annual major versions."),
    (6, "Enable SSO to OpenAIRE services PROVIDE, CONNECT and MONITOR with SURFconext",
     "80 institutions can log in with their SURFconext account.",
     "configuration", 9, 9, "9", "", ""),
]

# theme_no (0 = whole project), item, description, amount_eur, shared
BUDGET = [
    (1, "National symposium / conference / workshop",
     "National Symposium incl. General Assembly each year — € 23.000 × 4",
     92000, "no"),
    (1, "Travel and accommodation",
     "(International) travel and accommodation incl. conference visit", 3000, "yes"),
    (2, "Hardware",
     "Running monitor and reporting services — SURF Research Cloud VMs "
     "(€ 5.000 × 4 years)", 20000, "yes"),
    (2, "Work by third parties", "Development of the monitors for reporting",
     40000, "yes"),
    (2, "Travel and accommodation",
     "(International) travel and accommodation incl. conference visit", 3000, "yes"),
    (2, "Training and course",
     "Universities attending courses on metadata enrichment and full-text capture "
     "(e.g. JISC or OpenAIRE) — 15 persons × € 1.800", 27000, "yes"),
    (2, "Work by third parties",
     "Legal support, e.g. report on licencing and copyright of metadata and full text "
     "— € 300/h × 70 hours", 21000, "yes"),
    (3, "Hardware",
     "Running monitor and reporting services — SURF Research Cloud VMs "
     "(€ 5.000 × 4 years)", 20000, "yes"),
    (3, "Work by third parties", "Development of the monitors for reporting",
     40000, "yes"),
    (3, "Training and course",
     "Universities attending courses on metadata enrichment and full-text capture "
     "(e.g. JISC or OpenAIRE) — 15 persons × € 1.800", 27000, "yes"),
    (3, "Work by third parties",
     "Legal support, e.g. report on licencing and copyright of metadata and full text "
     "— € 300/h × 70 hours", 21000, "yes"),
    (4, "Hardware",
     "Running monitor and reporting services — SURF Research Cloud VMs "
     "(€ 5.000 × 4 years)", 20000, "yes"),
    (4, "Work by third parties", "Development of the monitors for reporting",
     40000, "yes"),
    (4, "Work by third parties",
     "Legal support, e.g. report on licencing and copyright of metadata and full text "
     "— € 300/h × 70 hours", 21000, "yes"),
    (5, "Hardware",
     "Running monitor and reporting services — SURF Research Cloud VMs "
     "(€ 5.000 × 4 years)", 20000, "yes"),
    (5, "Work by third parties", "Development of the monitors for reporting",
     40000, "yes"),
    (6, "Work by third parties", "Further development of the NL Research Portal",
     50000, "no"),
]

# theme_no, question, links -- exploratory notes carried over from slide 9.
OPEN_QUESTIONS = [
    (5, "What about discovery regulation for AI bots, adding AI usage licence markers?",
     "RSL: Really Simple Licensing :: https://rslstandard.org/"),
    (5, "What about discovery by proactive content push with Event Notifications in "
        "Value-Adding Networks?",
     "Event Notifications :: https://www.eventnotifications.net/"),
    (5, "What about discovery by signposting?",
     "Metadata Resources Pattern — Signposting the Scholarly Web :: "
     "https://signposting.org/patterns/metadata_resources/"),
]
