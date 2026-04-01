# Real Estate Data Scraper

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![BeautifulSoup](https://img.shields.io/badge/BeautifulSoup-4.x-59b256?style=for-the-badge)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-217346?style=for-the-badge&logo=microsoft-excel)
![Requests](https://img.shields.io/badge/Requests-2.x-FF6B35?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

*An automated property listing extraction engine with 
intelligent pagination handling and structured Excel export.*

</div>

---

## Overview

Real Estate Data Scraper is a purpose-built data 
extraction pipeline engineered to systematically 
harvest property listing data from paginated web 
sources. It abstracts the complexity of multi-page 
traversal, field normalization, and structured 
persistence into a single zero-configuration script.

The system is architected around a sequential 
crawl-parse-export pattern, implementing polite 
scraping conventions through inter-request delays 
to minimize server load and reduce detection risk.

---

## Architecture
```
real-estate-scraper/
│
├── real_estate_scraper.py     # Core extraction engine
│   ├── fetch_page()           # HTTP client layer
│   ├── parse_listings()       # DOM parsing & field extraction
│   ├── normalize_fields()     # Data cleaning & transformation
│   ├── paginate()             # Multi-page traversal controller
│   └── export_excel()         # Structured XLSX persistence
│
└── real_estate_data.xlsx      # Generated output artifact
```

---

## Technical Design

### Crawl-Parse-Export Pipeline
```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  HTTP Layer │────▶│  Parse Layer │────▶│ Export Layer│
│             │     │              │     │             │
│ requests    │     │BeautifulSoup │     │  openpyxl   │
│ User-Agent  │     │DOM traversal │     │ XLSX output │
│ time.sleep  │     │Field extract │     │Auto-fit cols│
└─────────────┘     └──────────────┘     └─────────────┘
        │                                       ▲
        │         ┌──────────────┐              │
        └────────▶│  Pagination  │──────────────┘
                  │  Controller  │
                  │ Auto-detect  │
                  │  next page   │
                  └──────────────┘
```

### Field Extraction & Normalization

Each listing is parsed for the following fields 
with graceful fallback to `"N/A"` on absence:

| Field | Source | Transformation |
|-------|--------|----------------|
| Property Name | `<article>` title element | Strip whitespace |
| Price | Price container | Strip currency symbols |
| Rating | `word-X` CSS class | Word → Integer mapping |
| Status | Availability tag | Normalize to In Stock / Out of Stock |

**Rating Normalization Map:**
```python
rating_map = {
    "One": 1, "Two": 2, "Three": 3,
    "Four": 4, "Five": 5
}
```

### Pagination Controller

The crawler implements automatic next-page detection 
by querying the DOM for pagination elements on each 
response cycle, terminating cleanly when no 
subsequent page reference is found — requiring 
zero hardcoded page counts.

### Polite Crawling
```python
time.sleep(1)  # 1-second inter-request delay
```
Implements respectful crawling conventions per 
standard web scraping ethics guidelines.

### Excel Output Specification

Generated `.xlsx` artifact features:
- **Bold, colored header row** (#4472C4 fill, 
  white font)
- **Auto-adjusted column widths** based on 
  content length
- **Frozen header row** for large dataset navigation
- **Row alternation** for readability (roadmap)

---

## Installation

**Prerequisites:** Python 3.8+
```bash
# Clone repository
git clone https://github.com/codewithmahfuz/real-estate-scraper.git
cd real-estate-scraper

# Install dependencies
pip install requests beautifulsoup4 openpyxl

# Run
python real_estate_scraper.py
```

---

## Usage
```bash
python real_estate_scraper.py
```
```
🔄 Scraping page 1...
🔄 Scraping page 2...
🔄 Scraping page 3...
...
🔄 Scraping page 50...

✅ Complete — 1000 listings saved to real_estate_data.xlsx
```

---

## Output Schema

| Column | Type | Description |
|--------|------|-------------|
| Property Name | String | Listing title |
| Price | String | Listed price |
| Rating | Integer | 1–5 numeric scale |
| Status | String | Availability status |

---

## Limitations & Known Constraints

- **JavaScript-rendered listings:** Single-page 
  applications (React, Vue, Angular) that inject 
  listing content client-side are outside current 
  scope. Playwright integration is planned.

- **Authentication walls:** Login-gated listing 
  platforms require session cookie injection — 
  not implemented in current version.

- **Dynamic pagination:** Sites using infinite 
  scroll rather than traditional pagination 
  require a different traversal strategy.

---

## Roadmap

- [ ] Async concurrent fetching via `aiohttp`
- [ ] Playwright integration for JS-rendered pages
- [ ] Configurable target URL via CLI arguments
- [ ] SQLite storage backend
- [ ] Proxy rotation & session management
- [ ] Docker containerization for portable deployment

---

## Ethical Usage

This tool is developed strictly for educational 
and research purposes using publicly available, 
scraping-friendly demo data. The author assumes 
no liability for misuse. Users are solely 
responsible for compliance with target platforms' 
Terms of Service and applicable data protection 
law including GDPR and CCPA.

---

## License

Distributed under the MIT License.
See `LICENSE` for details.

---

<div align="center">

**Built by [Mahfuz](https://github.com/codewithmahfuz)**  
Freelance Data Extraction & Automation Specialist

</div>
