# 🛠️ Bid Compare Tool

A powerful web application for comparing construction bids from multiple suppliers. Built with FastAPI (Python) and Vue.js, featuring advanced statistical analysis including Z-score calculations.

Excited to share my first app release! 🛠️

  I built this bid comparison tool to solve a real problem in the construction industry - comparing multiple offers
   efficiently. The app reads NS3459 XML files (Norwegian standard) and Excel/CSV, then provides statistical
  analysis to help you identify the best bid at a glance.

  Features Z-score analysis, chapter summaries, and automated Excel reports. Built with Python (FastAPI) and
  Vue.js, fully dockerized and open source.

  Available now on GitHub - would love your feedback!

## ✨ Features

- **Multi-format Support**: Upload bids in CSV, Excel (.xlsx/.xls), or NS3459 XML format
- **Automatic Comparison**: Side-by-side comparison of bids from different suppliers
- **Statistical Analysis**:
  - Z-score calculation for each bid (requires 3+ bids)
  - Standard deviation and mean calculations
  - Spread percentage analysis
- **Chapter Summaries**: Group and analyze bids by chapter/category
- **Excel Export**: Download formatted comparison reports
- **Visual Indicators**: Color-coded badges showing bid performance
- **Responsive Design**: Works on desktop, tablet, and mobile

## 📊 What is Z-score?

Z-score is a statistical measure that helps you evaluate how far each bid deviates from the average:

- **Negative Z-score** (e.g., -1.5): Lower than average = **Better** (cheaper) ✅
- **Z-score near 0**: Close to average = Neutral
- **Positive Z-score** (e.g., +1.5): Higher than average = **Worse** (more expensive) ❌

The provider with the **lowest total Z-score** across all items is the most consistently competitive bidder.

## 🚀 Quick Start

### Using Docker (Recommended)

1. Clone the repository:
```bash
git clone <repository-url>
cd BidCompareStandalone
```

2. (Optional) Configure ports if defaults are in use:
```bash
cp .env.example .env
# Edit .env and change FRONTEND_PORT and/or BACKEND_PORT
```

3. Start the application:
```bash
docker compose up -d
```

4. Open your browser and navigate to:
```
http://localhost        # Default (port 80)
# or http://localhost:YOUR_PORT if you changed FRONTEND_PORT in .env
```

**Default ports:**
- Frontend: http://localhost (port 80)
- Backend API: http://localhost:8000

**Port conflicts?** If ports 80 or 8000 are already in use on your system:
1. Copy `.env.example` to `.env`
2. Edit `.env` and set different ports:
   ```
   FRONTEND_PORT=3001
   BACKEND_PORT=8001
   ```
3. Run `docker compose up -d`

### Manual Setup

#### Backend

1. Navigate to the backend directory:
```bash
cd backend
```

2. Create a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the server:
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

#### Frontend

1. Navigate to the frontend directory:
```bash
cd frontend
```

2. Install dependencies:
```bash
npm install
```

3. Run the development server:
```bash
npm run dev
```

4. Open your browser to `http://localhost:3000`

## 📖 Usage Guide

### Step 1: Upload Bid Files

Click the file input and select one or more bid files. Supported formats:
- **CSV**: Comma or semicolon-separated values
- **Excel**: .xlsx or .xls files
- **NS3459 XML**: Norwegian construction standard format

### Step 2: Run Comparison

Click "Kjør sammenligning" (Run Comparison) to analyze the bids.

### Step 3: Review Results

The tool generates several views:

#### Summary (Oppsummering)
- Total number of items
- Lowest total bid (winner)
- Best Z-score provider (if 3+ bids)
- Individual bid totals with Z-score badges

#### Chapter Summary (Kapitteloppsummering)
- Bids grouped by chapter/category
- Best bidder per chapter
- Price spread analysis

#### Normalized Bids
- Detailed view of each bid
- All items with quantities and prices
- Option items shown in parentheses

#### Comparison Matrix (Sammenligning per postnr)
- Side-by-side comparison of all bids
- Statistical metrics (mean, std deviation, std %)
- Z-score columns (toggle "Vis Z-score" to show)

### Step 4: Export Results

Download comparison reports in Excel format:
- **Full Comparison**: Complete analysis with all bids
- **Matrix Excel**: Formatted comparison matrix
- **Chapter Summary**: Chapter-by-chapter analysis

## 🎨 Understanding the Color Badges

Z-score badges help you quickly identify bid performance:

- 🟢 **Green badge** (Z < -0.5): Significantly cheaper than average
- 🟡 **Yellow badge** (-0.5 ≤ Z ≤ 0.5): Close to average
- 🔴 **Red badge** (Z > 0.5): Significantly more expensive than average

## 🏗️ Project Structure

```
BidCompareStandalone/
├── backend/
│   ├── app/
│   │   └── main.py           # FastAPI application
│   ├── Dockerfile
│   └── requirements.txt
├── frontend/
│   ├── src/
│   │   ├── App.vue           # Main Vue component
│   │   ├── main.js           # Vue entry point
│   │   └── styles.css        # Application styles
│   ├── index.html
│   ├── package.json
│   ├── vite.config.js
│   ├── Dockerfile
│   └── nginx.conf
├── bid_compare_cli.py        # Command-line interface
├── docker-compose.yml
├── .env.example
├── .gitignore
├── LICENSE
└── README.md
```

## 🔧 API Documentation

### POST /api/bid-compare

Upload and compare bid files.

**Request:**
- Content-Type: `multipart/form-data`
- Body: One or more files

**Response:**
```json
{
  "normalized": { ... },
  "matrix": { ... },
  "chapters": { ... },
  "summary": {
    "totals": { "Provider A": 1500000, "Provider B": 1450000 },
    "option_totals": { "Provider A": 50000, "Provider B": 45000 },
    "winner": { "name": "Provider B", "total": 1450000 },
    "post_count": 150
  },
  "excel": "base64_encoded_data",
  "matrix_excel": "base64_encoded_data",
  "chapters_excel": "base64_encoded_data",
  "errors": []
}
```

### GET /health

Health check endpoint.

## 🛠️ Technology Stack

**Backend:**
- FastAPI - Modern Python web framework
- Pandas - Data analysis and manipulation
- NumPy - Numerical computing
- OpenPyXL - Excel file generation

**Frontend:**
- Vue 3 - Progressive JavaScript framework
- Vite - Next-generation frontend tooling

**Deployment:**
- Docker - Containerization
- Docker Compose - Multi-container orchestration
- Nginx - Web server and reverse proxy

## 💻 Command Line Interface (CLI)

For users who prefer working in the terminal or need to automate bid comparisons, a CLI version is available:

### Installation

```bash
cd /path/to/BidCompareStandalone
pip install -r backend/requirements.txt
```

### Usage

**Basic comparison (creates sammenligning.xlsx):**
```bash
python bid_compare_cli.py tilbud1.xlsx tilbud2.xml tilbud3.csv
```

**Custom output filename:**
```bash
python bid_compare_cli.py -o min_rapport.xlsx tilbud*.xml
```

**Verbose mode (show chapter breakdown):**
```bash
python bid_compare_cli.py -v tilbud1.xlsx tilbud2.csv
```

**Help:**
```bash
python bid_compare_cli.py --help
```

### CLI Output

The CLI shows **key metrics** in the terminal:
- Summary with all bids sorted by price
- Winner with lowest total
- Z-score analysis (if 3+ bids)
- Excel file automatically saved with **full detailed analysis**

With `-v` flag, also shows chapter-by-chapter breakdown in terminal.

### Example Output

```
Laster tilbud...
  ✓ tilbud1.xlsx -> Leverandør A
  ✓ tilbud2.xml -> Leverandør B
  ✓ tilbud3.csv -> Leverandør C

================================================================================
OPPSUMMERING
================================================================================

Antall tilbydere: 3
Antall poster: 145

TILBUD (eksklusive opsjoner):
--------------------------------------------------------------------------------
  Leverandør B                              kr     1,450,000.00 (+ kr 45,000.00 i opsjoner)
  Leverandør A                              kr     1,500,000.00 (+ kr 50,000.00 i opsjoner)
  Leverandør C                              kr     1,550,000.00

🏆 VINNER: Leverandør B                      kr     1,450,000.00

Z-SCORE TOTALER (lavere = bedre):
--------------------------------------------------------------------------------
  ✅ Leverandør B                                    -12.45
     Leverandør A                                      0.34
  ⚠️  Leverandør C                                     12.11

✅ Excel-rapport lagret: sammenligning.xlsx
```

The Excel file contains complete analysis with all posts, comparison matrix, and chapter summaries.

## 📝 File Format Requirements

### NS3459 XML Files (Anbefalt / Recommended)

**NS3459** er den norske standarden for elektronisk utveksling av tilbudsdata i byggebransjen. Dette er det anbefalte formatet.

The tool automatically extracts:
- Company name from metadata
- All bid items with prices
- Chapter structure (kapittel)
- Option items (opsjoner)
- NS codes and specifications

✅ **Best practice:** Bruk NS3459 XML-filer for mest nøyaktig sammenligning.

### CSV/Excel Files (Alternativ)

Hvis du ikke har NS3459 XML, kan du bruke CSV eller Excel med følgende kolonner:

**Påkrevde kolonner (case-insensitive):**
- `postnr` - Postnummer (f.eks. "01.10", "02.15")
- `beskrivelse` eller `description` - Beskrivelse av posten
- `enhet` eller `unit` - Enhet (stk, m2, m3, etc.)
- `mengde` eller `qty` - Mengde/antall

**Priser (minst én av disse):**
- `pris` eller `unit_price` - Enhetspris
- `sum` eller `sum_amount` - Totalsum (hvis bare sum er oppgitt, brukes denne)

**Valgfrie kolonner:**
- `kode`, `nskode`, `ns_code`, eller `code` - NS-kode

**Eksempel på Excel-fil:**

| postnr | beskrivelse           | enhet | mengde | pris    | sum      |
|--------|-----------------------|-------|--------|---------|----------|
| 01.10  | Grunnarbeid          | m2    | 100    | 150.00  | 15000.00 |
| 01.20  | Fundamentering       | m3    | 25     | 2500.00 | 62500.00 |
| 02.10  | Murverk              | m2    | 200    | 850.00  | 170000.00|

**Tips:**
- Kolonnenavnene kan være både norske og engelske
- Store/små bokstaver spiller ingen rolle
- Kapittel utledes automatisk fra postnummer (f.eks. "01" fra "01.10")

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built for the Norwegian construction industry
- Supports NS3459 standard for bid exchange
- Designed for analyzing tenders in accordance with Norwegian building practices

## 📧 Support

For issues, questions, or suggestions, please open an issue on GitHub.

---

Made with ❤️ for construction professionals
