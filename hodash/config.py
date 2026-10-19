"""Everything you might want to tweak lives here (paths, file patterns, mappings)."""
import os
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
# Folder that is watched for sales files. Default = the "Sales Review" folder this app sits in.
_DEFAULT_DATA = APP_DIR.parent / "Data" if (APP_DIR.parent / "Data").is_dir() else APP_DIR.parent
DATA_DIR = Path(os.environ.get("SALES_DATA_DIR", _DEFAULT_DATA))
CACHE_DIR = APP_DIR / ".cache"
TARGET_PATTERN = "target*"                 # Target workbook in DATA_DIR (newest match is used)
TARGET_SHEET = "FY2526 Base Target"        # sheet holding the current-year target blocks
HO_TEMPLATE_GLOB = "HO PPT format*TEMPLATE*.pptx"   # looked up in DATA_DIR

# File-name patterns (case-insensitive) that are picked up from DATA_DIR.
# Drop "Primary Sales Oct26.xlsx" / "Secondary Sales Oct26.xlsx" etc. and they are merged automatically.
PATTERNS = {
    "primary": "primary*",
    "secondary": "secondary*",
    "material": "material master*",
    "pack": "pack master*",
}
EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".csv"}

# Secondary rows whose SKU name contains one of these are dropped (not lubes: ClearBlue = DEF).
EXCLUDE_SKU_PATTERNS = ["CLEARBLUE"]

# Secondary 'Customer Group' -> reporting channel (matches the HO deck: ServoXpert is counted in Bazaar,
# 'Dealer' = IOC retail outlets = RO).
CHANNEL_MAP = {
    "BAZAAR": "Bazaar", "SERVOXPERT": "Bazaar", "OTHERS": "Bazaar",
    "DEALER": "RO", "NAYARA RO": "RO",
    "OEM": "OEM", "FLEET": "Fleet",
}
CHANNEL_ORDER = ["Bazaar", "RO", "OEM", "Fleet"]

# Display names per stockist code (falls back to the cleaned name in the file).
STOCKIST_ALIASES = {
    199950: "Doshi Enterprises",
    200749: "AMCO Distributor [SSA]",
    262098: "Madhuban Enterprises",
}

# --- HO-template groupings -------------------------------------------------------------
# Pack category (Pack Master) -> HO slide "SKU Group"
PACK_GROUPS = {
    "Less than 250ml": "Pouch <250 ml",
    "250ml-1 ltr": "250 ml to 1 ltr",
    "1 ltr to <3 & 3.5  ltr": "> 1 ltr to 6 ltr",
    "3 & 3.5 ltr": "> 1 ltr to 6 ltr",
    "4  To 5 ltr": "> 1 ltr to 6 ltr",
    "> 5 ltr to < 10 ltr": "> 6 ltr and < 15 ltr",
    "> 10 ltr to < 20 ltr": "> 6 ltr and < 15 ltr",
    "20 ltr": "> 15 ltr to 25 ltr",
    "26 ltr": "26 ltr",
    "50 ltr": "50 ltr",
    "210, 208, 205 , 200 ltr": "210, 208,205,200 ltr",
    "182, 175 kg, 200 kg": "182, 175 kg, 200 kg",
}
PACK_GROUP_ORDER = ["Pouch <250 ml", "250 ml to 1 ltr", "> 1 ltr to 6 ltr", "> 6 ltr and < 15 ltr",
                    "> 15 ltr to 25 ltr", "26 ltr", "50 ltr", "210, 208,205,200 ltr", "182, 175 kg, 200 kg"]

PRODUCT_GROUP_ORDER = ["4T Oil", "PCMO", "DEO", "Gear Oil", "Agri", "Grease", "Speciality", "2T Oil",
                       "Hydrashakti 68", "System 68", "Other Industrial", "Genuine Oil"]

FY_START_MONTH = 4

# --- scheme calculator ---------------------------------------------------------------
SCHEME_DIR = Path(os.environ.get("SCHEME_DIR", APP_DIR.parent.parent / "Lube Scheme Circulars"))
RULES_FILE = SCHEME_DIR / "Scheme Rules.xlsx"      # editable rule book read by the incentive calculator
