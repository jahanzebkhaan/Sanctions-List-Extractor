# 🌐 Global Sanctions List Extractor

A Python tool that automatically downloads, parses, and consolidates **four major international sanctions lists** into a single, clean Excel workbook — ready for compliance screening and analysis.

---

## 📋 Data Sources

| # | List | Authority | Format | Update Frequency |
|---|------|-----------|--------|-----------------|
| 1 | **OFAC SDN** | US Treasury | XML | Daily |
| 2 | **UK HMT OFSI** | His Majesty's Treasury | CSV | Weekly |
| 3 | **EU FSF** | European Commission | CSV | Daily |
| 4 | **UNSC** | UN Security Council | XML | As needed |

---

## 📊 Output

A formatted Excel file (`sanctions_output.xlsx`) with **6 sheets**:

| Sheet | Contents |
|-------|----------|
| **All Records** | Combined view — all sources, all 32 columns |
| **Summary** | Record counts, fields captured, generation timestamp |
| **OFAC** | OFAC-only records with source-specific columns |
| **UK_HMT** | UK HMT-only records |
| **EU** | EU FSF-only records |
| **UNSC** | UNSC-only records |

### Columns Extracted (32 total)

```
Source                    Reference / UID           Sanctions Programme
Last Name                 First Name                Second / Middle Name
Third Name                Fourth Name               Title
Designation / Position    Gender                    Good Quality AKA
Low / Weak Quality AKA    Date of Birth             Place of Birth
Country of Birth          Nationality 1             Nationality 2
Passport Number           Passport Country          Passport Expiry
National ID Number        National ID Type          National ID Country
Address                   City                      Country
Postal Code               Listed On                 Last Updated
Group Status              Remarks / Other Info
```

---

## 🚀 Setup

**Requirements:** Python 3.8+

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/sanctions-list-extractor.git
cd sanctions-list-extractor

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python sanctions_extractor.py
```

That's it. The script downloads all four lists, parses them, and saves the Excel file in the same directory.

---

## 📁 Project Structure

```
sanctions-list-extractor/
│
├── sanctions_extractor.py   # Main script
├── requirements.txt         # Python dependencies
├── README.md                # This file
├── .gitignore               # Excludes cache and output files
└── sanctions_cache/         # Auto-created — cached downloads (gitignored)
```

---

## ⚙️ How It Works

```
[Download]          [Parse]             [Export]
OFAC XML    ──►
UK HMT CSV  ──►  Normalise into    ──►  Excel (.xlsx)
EU CSV      ──►  32 master columns      6 sheets
UNSC XML    ──►
```

1. **Download** — Each list is fetched from its official source URL and cached locally. Re-runs use the cache (delete `sanctions_cache/` to force a fresh download).
2. **Parse** — Source-specific parsers extract all available fields and map them to a common 32-column schema. Illegal characters (e.g. null bytes from Arabic script) are stripped automatically.
3. **Export** — Records are written to a colour-coded, filtered, freeze-paned Excel file grouped by source.

---

## 🔄 Keeping Lists Fresh

The cache folder stores downloaded files so repeated runs are instant. To refresh:

```bash
# Delete cache to re-download all lists
rm -rf sanctions_cache/

# Or delete a specific list
rm sanctions_cache/ofac_sdn.xml
```

---

## 📦 Dependencies

```
requests
openpyxl
```

Install with:
```bash
pip install -r requirements.txt
```

---

## ⚠️ Disclaimer

This tool is intended for **compliance research and analytical purposes only**. Always verify against official sources before making compliance decisions. The authors are not responsible for decisions made based on this data.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you'd like to change.
