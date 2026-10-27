"""Pincode -> Area -> Zone seed list and RO location seeds for the area-wise secondary report.

These are only the *starting* values written into `Data/Area Master.xlsx` the first time the report
runs. After that the workbook is the source of truth: edit it there, not here.
"""

# Zone order used in every report table (top to bottom = south to north).
ZONE_ORDER = [
    "South Mumbai", "Central Mumbai", "Eastern Suburbs", "Western Suburbs - Bandra to Andheri",
    "Western Suburbs - Jogeshwari to Dahisar", "Thane & Navi Mumbai", "Mira-Bhayandar", "Vasai-Virar",
    "Palghar-Boisar", "Dahanu-Talasari", "Wada-Jawhar-Vikramgad", "Unmapped",
]

_SM, _CM, _ES = "South Mumbai", "Central Mumbai", "Eastern Suburbs"
_WA, _WD = "Western Suburbs - Bandra to Andheri", "Western Suburbs - Jogeshwari to Dahisar"
_TN, _MB, _VV = "Thane & Navi Mumbai", "Mira-Bhayandar", "Vasai-Virar"
_PB, _DT, _WJ = "Palghar-Boisar", "Dahanu-Talasari", "Wada-Jawhar-Vikramgad"

# pincode: (Area, Zone). Area = post-office locality, Zone = sales territory.
PINCODES = {
    "400001": ("Fort / Ballard Estate", _SM), "400002": ("Kalbadevi", _SM), "400003": ("Masjid Bunder", _SM),
    "400004": ("Girgaon", _SM), "400005": ("Colaba", _SM), "400006": ("Malabar Hill", _SM),
    "400007": ("Grant Road", _SM), "400008": ("Mumbai Central", _SM), "400009": ("Dongri / Chinchbunder", _SM),
    "400010": ("Mazgaon", _SM), "400011": ("Jacob Circle / Agripada", _SM), "400020": ("Marine Lines / Churchgate", _SM),
    "400021": ("Nariman Point", _SM), "400023": ("Fort", _SM), "400026": ("Cumballa Hill", _SM),
    "400027": ("Byculla", _SM), "400032": ("Mantralaya", _SM), "400034": ("Tardeo", _SM),
    "400035": ("Malabar Hill", _SM), "400036": ("Peddar Road", _SM),
    "400012": ("Parel", _CM), "400013": ("Lower Parel", _CM), "400014": ("Dadar East", _CM),
    "400015": ("Sewri", _CM), "400016": ("Mahim", _CM), "400017": ("Dharavi", _CM), "400018": ("Worli", _CM),
    "400019": ("Matunga", _CM), "400022": ("Sion", _CM), "400025": ("Prabhadevi", _CM), "400028": ("Dadar West", _CM),
    "400030": ("Worli", _CM), "400031": ("Wadala", _CM), "400033": ("Cotton Green / Kalachowki", _CM),
    "400037": ("Antop Hill / Wadala East", _CM),
    "400024": ("Kurla (Nehru Nagar)", _ES), "400042": ("Bhandup East", _ES), "400043": ("Mankhurd / Govandi", _ES),
    "400070": ("Kurla West", _ES), "400071": ("Chembur", _ES), "400072": ("Sakinaka", _ES), "400074": ("Chembur", _ES),
    "400075": ("Ghatkopar East (Pant Nagar)", _ES), "400076": ("Powai", _ES), "400077": ("Ghatkopar East (Rajawadi)", _ES),
    "400078": ("Bhandup West", _ES), "400079": ("Vikhroli", _ES), "400080": ("Mulund West", _ES),
    "400081": ("Mulund East", _ES), "400082": ("Mulund", _ES), "400083": ("Vikhroli East", _ES),
    "400084": ("Ghatkopar West (Barve Nagar)", _ES), "400085": ("Trombay", _ES), "400086": ("Ghatkopar West", _ES),
    "400087": ("Powai", _ES), "400088": ("Govandi / Trombay", _ES), "400089": ("Chembur (Tilak Nagar)", _ES),
    "400094": ("Anushakti Nagar", _ES),
    "400029": ("Santacruz East", _WA), "400049": ("Juhu", _WA), "400050": ("Bandra West", _WA),
    "400051": ("Bandra East / BKC", _WA), "400052": ("Khar", _WA), "400053": ("Andheri West", _WA),
    "400054": ("Santacruz West", _WA), "400055": ("Santacruz East", _WA), "400056": ("Vile Parle West", _WA),
    "400057": ("Vile Parle East", _WA), "400058": ("Andheri West", _WA), "400059": ("Marol / Andheri East", _WA),
    "400061": ("Versova", _WA), "400069": ("Andheri East", _WA), "400093": ("Chakala MIDC", _WA),
    "400096": ("SEEPZ", _WA), "400098": ("Kalina / Vidyanagari", _WA), "400099": ("Sahar / Airport", _WA),
    "400060": ("Jogeshwari East", _WD), "400062": ("Goregaon West", _WD), "400063": ("Goregaon East", _WD),
    "400064": ("Malad West", _WD), "400065": ("Aarey Colony", _WD), "400066": ("Borivali East", _WD),
    "400067": ("Kandivali West", _WD), "400068": ("Dahisar", _WD), "400091": ("Borivali West (Gorai)", _WD),
    "400092": ("Borivali West", _WD), "400095": ("Malad West (Malvani)", _WD), "400097": ("Malad East", _WD),
    "400101": ("Kandivali East", _WD), "400102": ("Jogeshwari West / Oshiwara", _WD),
    "400103": ("Borivali West (Mandapeshwar)", _WD), "400104": ("Goregaon West (Bangur Nagar)", _WD),
    **{f"4006{n:02d}": ("Thane", _TN) for n in range(1, 16)},
    **{f"4007{n:02d}": ("Navi Mumbai", _TN) for n in range(1, 11)},
    "401101": ("Bhayandar West", _MB), "401104": ("Mira Road (Kashimira)", _MB), "401105": ("Bhayandar East", _MB),
    "401106": ("Uttan / Bhayandar", _MB), "401107": ("Mira Road", _MB), "401701": ("Mira Road (check pincode)", _MB),
    "401201": ("Vasai West", _VV), "401202": ("Vasai Road", _VV), "401203": ("Nallasopara", _VV),
    "401204": ("Khanivade", _VV), "401205": ("Vasai (Rural)", _VV), "401206": ("Vasai (Rural)", _VV),
    "401207": ("Vasai West", _VV), "401208": ("Vasai East (Waliv / Gokhivare)", _VV),
    "401209": ("Nallasopara East", _VV), "401210": ("Vasai West", _VV), "401301": ("Virar", _VV),
    "401302": ("Virar West (Arnala / Agashi)", _VV), "401303": ("Virar East", _VV), "401305": ("Virar", _VV),
    "401102": ("Saphale", _PB), "401103": ("Tarapur / Vangaon", _PB), "401401": ("Kelwe", _PB),
    "401402": ("Khadkoli", _PB), "401403": ("Manor", _PB), "401404": ("Palghar", _PB), "401501": ("Boisar", _PB),
    "401502": ("Boisar West (Tarapur)", _PB), "401503": ("Chinchani", _PB), "401504": ("Boisar", _PB),
    "401506": ("Boisar MIDC", _PB),
    "401601": ("Dahanu Road", _DT), "401602": ("Dahanu", _DT), "401604": ("Dahanu", _DT),
    "401606": ("Talasari", _DT), "401607": ("Kasa", _DT),
    "401603": ("Jawhar", _WJ), "401605": ("Vikramgad", _WJ), "421302": ("Wada (Ambadi)", _WJ),
    "421303": ("Wada", _WJ), "421312": ("Wada (Kudus)", _WJ),
}

# RO (Dealer) customers are not in the Lubes Customer Master. Pincodes below come from the area each
# RO was tagged with in the user's FY25-26 sheet "Taluka wise sales of SSR doshi.xlsx".
RO_SEED = {
    "0000125769": ("400011", "GUJARAT SERVICE CENTRE"), "0000125773": ("400053", "SAHARA HOSPITALITY LTD."),
    "0000125825": ("400069", "ROYAL SERVICE"), "0000125795": ("400050", "KINI'S CAUSEWAY"),
    "0000196265": ("400051", "3-STAR FUEL STATION"), "0000280430": ("400051", "COCO BKC"),
    "0000310216": ("400051", "COCO BKC"), "0000356194": ("400051", "COCO BKC"),
    "0000108268": ("400067", "CHAMUNDI PETROLEUM"), "0000308413": ("400026", "COCO TOP GEAR"),
    "0000342167": ("400026", "COCO TOP GEAR"), "0000370551": ("400026", "COCO TOPGEAR"),
    "0000230873": ("400026", "SAI DURGA SERVICES"), "0000354411": ("400028", "COCO PRABHADEVI"),
    "0000125846": ("400068", "SATYASHANTI AUTO SERVICE - DAHISAR"), "0000192105": ("400004", "MUMBAI PETROL SUPPLY CO"),
    "0000192096": ("400063", "FORT MOTOR EQUIPMENT CO"), "0000293033": ("400063", "FORT MOTORS EQUIPMENT CO.-ADHOC COC"),
    "0000125772": ("400063", "GOREGAON AUTO CENTRE"), "0000125857": ("400063", "TAXIMANS SERVICES"),
    "0000192125": ("400060", "PUSHPAK AUTO & GAS STATION"), "0000125744": ("400049", "ACE ECO STATION"),
    "0000349074": ("400060", "ANURADHA GAS"), "0000125771": ("400067", "GARUDA AUTO SERVICE"),
    "0000125845": ("400067", "SHAH AUTOMOBILES"), "0000125842": ("400016", "SOVEREIGN SERVICE STATION"),
    "0000226293": ("400064", "GURUPRASAD SERVICES"), "0000358220": ("400020", "COCO III MARINE LINES"),
    "0000219959": ("400020", "COCO -III- MARINE LINES"), "0000223850": ("400102", "RAJENDRA SERVICES"),
    "0000125781": ("400099", "I CARPORT"), "0000312477": ("400099", "COCO SAHAR"),
    "0000125754": ("400072", "BON BON SERVICE CENTRE"), "0000164442": ("400057", "RAJ AUTO SERVICE"),
    "0000288263": ("400018", "SHREE KRISHNA AUTOMOBILES - ADHOC B"), "0000192108": ("400018", "SHRI KRISHNA AUTOMOBILES PVT.LTD."),
    "0000125838": ("400018", "SURAJ SANGHI SERVICE CENTRE"),
}
