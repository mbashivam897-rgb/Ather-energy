# Ather Energy Ltd — Integrated Financial Model

**File:** `Ather_Financial_Model.xlsx`

An institutional, equity-research-style financial model for **Ather Energy Limited**, built to mirror the
structure, colour coding and modelling philosophy of the supplied `United Spirits Excell.xlsx` template.

- **History:** FY2021 – FY2025 (from the annual reports in this repo)
- **Forecast:** FY2026E – FY2033E (8 years)
- **Currency:** INR millions unless stated
- **Basis:** Ather has no subsidiaries and files only **standalone** Ind AS statements, so standalone is used as the consolidated-equivalent (documented in the model).

## Workbook contents (14 sheets)

| Sheet | Purpose |
|---|---|
| Cover | Index, headline outputs, colour key |
| Scenarios | Base / Bull / Bear switch (volume growth & gross margin) |
| Assumptions | Every macro, industry and company driver, each with a **Source** and a plain-language **Reason** |
| Revenue Build | Driver-based revenue: units × ASP + accessories + software + charging + service |
| PL | Income statement |
| BS | Balance sheet |
| CS | Cash flow statement (fully articulated) |
| Working Capital Schedule | Receivables / inventory / payables and change in NWC |
| Working Schedules | PPE, ROU (lease) assets, intangibles, CWIP, D&A, borrowings, lease liability, finance cost, tax & carry-forward losses |
| Ratio Analysis | Profitability, liquidity, leverage, efficiency, per-share |
| DCF | FCFF DCF with CAPM-based WACC and a normalised terminal value |
| Error Checks | Automated integrity checks (all return PASS) |
| Company & Market Insights | Industry, competition, battery costs, policy, sources |
| Modelling Notes | Methodology, assumptions rationale, risks, limitations |

## Colour convention

- **Blue** = hardcoded input
- **Green** = link / reference to another sheet
- **Black** = in-sheet calculation

## Key design points

- **Fully linked & dynamic** — no hardcoded numbers inside forecast formulas; no circular references (interest uses opening balances).
- **Fully articulated** — every balance-sheet line is forecast from a driver and cash is the genuine output of the cash-flow statement, so the **balance sheet balances by accounting identity (no plug)**.
- **Scenario engine** — change the single cell `Scenarios!F4` (1 = Base, 2 = Bull, 3 = Bear) to flex the whole model.
- **Validated** — recalculated with an independent formula engine: balance sheet balances every year, all integrity checks PASS, and there are zero formula errors.

## Headline (Base case)

- FY2033E revenue ≈ INR 106,638m; EBITDA breakeven ≈ FY2030; approaches PAT breakeven ≈ FY2032–FY2033.
- Illustrative DCF value ≈ INR 77 / share (Bull ≈ 133, Bear ≈ 4) — conservative and highly sensitive to terminal margin/WACC; read alongside relative valuation.

## Reproducibility

- `model_generator.py` is the Python (openpyxl) script that builds the workbook.
- `SOURCE_DATA_AND_RESEARCH.md` traces the historical data and market research with sources.

*For analytical / educational use. Not investment advice.*
