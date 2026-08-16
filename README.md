# EBR Drill Analytics

**EBR Drill Analytics** is a Streamlit application for analyzing Sandvik iSURE® drilling round reports from El Brocal.

The app processes one or multiple PDF reports, identifies the drilling jumbo, extracts drilled-hole information, validates the extracted records against the **TIPOS DE BARRENO** summary, displays a boxplot for each cycle, and generates a consolidated Excel workbook.

> This is an internal analytical tool and is not an official Sandvik product.

## Main features

- Upload one or multiple iSURE® PDF reports.
- Automatic jumbo identification:
  - `125D114796` → `JUMB001`
  - `125D98943` → `JUMB002`
- Reads cycle metadata:
  - cycle number
  - start date/time
  - serial number
  - drilling plan
  - drilled metres
- Reads the **TIPOS DE BARRENO** summary automatically.
- Extracts drilled-hole detail by type:
  - Bottom
  - Easer
  - Cut
  - Contour
  - Reaming
  - Casing
- Identifies extra/non-planned holes (`E1`, `E2`, etc.).
- Calculates:
  - minimum
  - maximum
  - average
  - median
  - preliminary axial length
- Generates a boxplot for each drilling cycle.
- Reconciles expected vs. extracted holes.
- Exports a consolidated Excel workbook with:
  - `Resumen_Reportes`
  - `Resumen_Ciclos`
  - `Detalle_Barrenos`
  - `Validacion`
  - `Barrenos_Extra`

## Repository structure

```text
EBR-Drill/
├── app.py
├── procesador.py
├── requirements.txt
└── README.md
```

## Run locally

Create and activate a Python virtual environment, then install the dependencies:

```bash
python -m pip install -r requirements.txt
```

Run the Streamlit app:

```bash
streamlit run app.py
```

Streamlit will open the application in the browser.

## Deploy to Streamlit Community Cloud

1. Push these files to the GitHub repository.
2. Sign in to Streamlit Community Cloud.
3. Create a new app.
4. Select the GitHub repository:
   `mioqui/EBR-Drill`
5. Select the `main` branch.
6. Set the main file path to:
   `app.py`
7. Deploy the application.
8. Keep the app private while testing operational reports, subject to company information-security policies.

## Important validation rule

The application compares extracted holes with the counts reported in the PDF page **TIPOS DE BARRENO**.

If a cycle displays `REVISAR`, the extracted data should not be considered validated until the discrepancy is reviewed.

## Data handling

When deployed on Streamlit Community Cloud, uploaded PDFs are processed on the Streamlit-hosted server. Confirm that this is acceptable under your organization's information-security requirements before using operational or sensitive reports.

## Trademark note

Sandvik and iSURE® are trademarks/products of Sandvik.  
EBR Drill Analytics is an independent analytical application and is not affiliated with or endorsed by Sandvik.
