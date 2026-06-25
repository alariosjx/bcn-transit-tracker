# export/fiscal_context.py
# Editorial context and metric definitions — separated so journalists can update
# without touching pipeline code.
# Bay City News | Andres Jimenez Larios

FISCAL_CONTEXT = {
    "bart": (
        "BART faces a $376M deficit in FY2027. Emergency federal COVID relief funds are "
        "exhausted. A $590M state bridge loan (AB/SB 117) is active through a Nov 2026 "
        "regional ballot measure. Failure at the ballot would trigger severe service cuts."
    ),
    "muni": (
        "Muni faces a $307M deficit growing to $430M by 2030. COVID relief funds exhausted. "
        "Included in the $590M state bridge loan package. Dependent on Nov 2026 ballot measure "
        "for long-term stability."
    ),
    "smart": (
        "SMART is not in a fiscal crisis. It operates on a smaller budget funded primarily "
        "by Measure Q sales tax in Sonoma and Marin counties. Ridership has grown steadily "
        "since service began in 2017."
    ),
    "actransit": (
        "AC Transit faces a $130M structural deficit. Included in the $590M state bridge "
        "loan package. Dependent on Nov 2026 regional ballot measure for long-term stability."
    ),
    "caltrain": (
        "Caltrain completed its electrification project in 2024, boosting ridership significantly. "
        "Faces long-term funding uncertainty but not in immediate crisis. Dependent on "
        "GoPass corporate program and fare revenue for operating costs."
    ),
    "sfbayferry": (
        "WETA operates SF Bay Ferry. Ridership has grown significantly post-pandemic. Expanding service to new routes."
    ),
    "samtrans": (
        "SamTrans faces structural deficits. Included in regional funding discussions tied to Nov 2026 ballot."
    ),
    "vta": (
        "VTA faces a $100M+ deficit. Included in $590M state bridge loan package. Dependent on Nov 2026 ballot."
    ),
    "ggferry": (
        "Golden Gate Ferry is financially stable relative to bus operations. Ridership recovering well post-pandemic."
    ),
    "ggbus": (
        "Golden Gate Bus faces significant deficits. Transbay routes dependent on commuter ridership recovery."
    ),
    "marin": (
        "Marin Transit is a relatively small system. Financially stable with Measure A sales tax funding."
    ),
    "napa": (
        "NVTA operates on Measure T sales tax. Small system, not in fiscal crisis."
    ),
    "vallejo": (
        "Vallejo Transit is small and locally funded. Not in immediate crisis."
    ),
    "countyconnection": (
        "County Connection faces modest deficits. Contra Costa County sales tax funded."
    ),
    "westcat": (
        "WestCAT is a small system serving western Contra Costa. Not in fiscal crisis."
    ),
    "tridelta": (
        "Tri Delta Transit serves eastern Contra Costa. Small system, locally funded."
    ),
    "wheels": (
        "Wheels/LAVTA serves Livermore-Pleasanton-Dublin. Locally funded, not in crisis."
    ),
    "unioncity": (
        "Union City Transit is a small local system. Not in fiscal crisis."
    ),
    "alamedaferry": (
        "Alameda Ferry serves the Oakland-SF transbay corridor. Financially stable."
    ),
    "santarosa": (
        "Santa Rosa CityBus is locally funded. Not in fiscal crisis."
    ),
    "fairfield": (
        "Fairfield-Suisun Transit is a small system. Not in fiscal crisis."
    ),
    "vacaville": (
        "Vacaville City Coach is a small local system. Not in fiscal crisis."
    ),
    "petaluma": (
        "Petaluma Transit is a small system. Not in fiscal crisis."
    ),
    "mst": (
        "Monterey-Salinas Transit serves Monterey County. Not in immediate fiscal crisis."
    ),
    "santacruz": (
        "Santa Cruz Metro faces modest funding challenges. Locally funded."
    ),
    "sjrtd": (
        "San Joaquin RTD serves Stockton area. Faces funding challenges typical of smaller systems."
    ),
    "ace": (
        "ACE commuter rail connects Stockton to San Jose. Faces funding uncertainty post-pandemic."
    ),
}

GLOSSARY = [
    ("Fare Revenue",        "Total fares collected from riders in that calendar year (NTD)."),
    ("Operating Expenses",  "Total cost to run the system that year — staff, maintenance, administration (NTD)."),
    ("Fare Recovery Ratio", "Share of operating costs covered by fares. Pre-pandemic BART: ~72%. 2024: ~25%."),
    ("Cost Per Trip",       "Operating expenses divided by total annual boardings. Higher = less efficient."),
    ("YoY Change",          "Percent change vs the same month one year prior."),
    ("Recovery vs 2019",    "Ridership as a percent of the equivalent 2019 month — pandemic recovery indicator."),
    ("12-Month Avg",        "Average monthly ridership over the last 12 complete months."),
]
