# 🏠 Rental Bill Automation — AI-powered PDF → Excel

> Automated processing of rental and utility bills from PDF into structured Excel outputs.  
> Extracts key financial data using LLM-based document intelligence — reducing ~30 minutes of manual work to ~15 seconds.

🔗 [Live Demo](https://rental-bill-automation-production.up.railway.app/)

---

## The Problem

Every month, landlords who rent out their apartments receive a utility bill (ЕПД) in PDF format. Calculating the tenant's share requires:

1. Manually copying 15+ line items from the PDF into Excel
2. Adding electricity costs from a separate source (МосЭнергоСбыт)
3. Applying correction formulas (recalculations + debt/overpayment − paid)
4. Splitting totals between tenant and landlord

The PDF structure **changes month to month** — different columns appear and disappear, new service lines are added. A rigid parser breaks immediately.

**Solution:** LLM-based document intelligence that understands the table semantically, not positionally.

---

## How It Works

```
PDF utility bill
      │
      ▼
pypdf → raw text
      │
      ▼
Claude API (LLM)
- Identifies which columns are present this month
- Extracts: volume, recalculations, debt/overpayment, paid — per service line
- Returns structured JSON
      │
      ▼
Python business logic
- Maps service names → Excel rows (with fuzzy matching)
- Calculates: correction = recalculations + debt/overpayment − paid
- Writes Excel formulas for all calculated sections
      │
      ▼
openpyxl → populated Excel file
      │
      ▼
Download via browser
```

---

## Key Technical Decisions

**Semantic extraction over positional parsing**  
The PDF table structure changes monthly. Instead of reading column N, the LLM reads by header name and adapts automatically. This solved the core reliability problem.

**Separation of extraction and calculation**  
Early version asked the AI to calculate corrections itself — it got the sign wrong when `zadolzhennost` equalled `oplacheno` (they cancel out to zero, but AI returned a negative number). Fixed by having AI extract raw numbers only; Python does the arithmetic.

**Formula injection**  
Rather than relying on the template having pre-existing formulas, `app.py` writes all Excel formulas programmatically for the target column on every run. The template is a clean schema — no prior data required.

**Month-aware column detection**  
Column is determined by the month name extracted from the PDF period (e.g. "март 2026" → column G). 
Falls back to the first empty column if the month is unrecognised. An earlier version read `=EOMONTH()` header formulas which appeared as `None` in read-only mode — causing all data to land in January regardless of the actual month.

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend | Python 3.12, Flask |
| AI / LLM | Anthropic Claude API (claude-sonnet) |
| PDF parsing | pypdf |
| Excel | openpyxl |
| Frontend | Vanilla HTML / CSS / JS, drag-and-drop UI |
| Deployment | Railway |

---

## Local Setup

```bash
# 1. Clone and install
git clone https://github.com/Gzelle-cloud/rental-bill-automation.git
cd rental-bill-automation
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Mac/Linux
pip install -r requirements.txt

# 2. Set your Anthropic API key
cp .env.example .env
# Open .env and paste your key

# 3. Run
python app.py

# 4. Open http://localhost:5050
```

---

## Usage

### Try it live
🔗 [rental-bill-automation-production.up.railway.app](https://rental-bill-automation-production.up.railway.app/)

### Or run locally
After completing local setup above, open `http://localhost:5050`

1. Upload the monthly PDF utility bill (drag & drop)
2. Upload your current Excel workbook — optional, uses clean template if not provided
3. Enter the electricity amount from МосЭнергоСбыт personal account
4. Click **Обработать квитанцию**
5. Review the calculated totals — tenant amount and landlord amount shown instantly
6. Download the updated Excel — named automatically, e.g. `Квитанции_март_2026.xlsx`

Next month, upload that downloaded file as the Excel input to accumulate the full year.

---

## Excel Structure

| Section | Rows | Source |
|---------|------|--------|
| Housing service volumes | 5–13 | "Объём услуг" column from PDF |
| Utility service volumes | 16–21 | "Объём услуг" column from PDF |
| Housing corrections | 36–44 | recalc + debt − paid (from PDF) |
| Utility corrections | 68–73 | recalc + debt − paid (from PDF) |
| Calculated totals (2.1, 2.3, 3.1, 3.3, 4.x) | various | Excel formulas written by app |
| Electricity IPU | 90 | User input |
| Tenant total | 92 | Excel formula |

---

## Project Context

Built as a personal automation tool and portfolio project demonstrating:
- Applied LLM integration in a real-world business workflow
- Document intelligence with variable/unpredictable schema
- End-to-end product thinking: problem definition → requirements → iterative testing → working solution
- Prompt engineering for structured data extraction from unstructured documents
- Originally built and tested on real-world municipal utility bills (ЖКХ), demonstrating robustness to semi-structured and evolving document formats

---

## Extensibility

The system is designed with separation of concerns between extraction, schema definition, and business logic, enabling adaptation to other semi-structured financial documents.

- Extraction layer is prompt-driven and can be adjusted for different PDF layouts
- Data schema is explicitly defined and reusable across document types
- Business logic is independent of document structure
