# EBR Drill Analytics

**EBR Drill Analytics** is a Streamlit application for analyzing drilling round data from underground jumbo drilling equipment.

The application processes one or multiple PDF and ZDA files, identifies drilling equipment and cycles, extracts and validates hole-level data, and provides consolidated analysis of drilling performance.

Key capabilities include drilling automation analysis by jumbo and boom, drilled-length analysis by hole type, cycle-time monitoring, drilling-round classification, spatial reconstruction of drilling patterns from ZDA data, and export of consolidated results to Excel.

> This is an internal analytical tool intended to support operational analysis and review.

## Main features

- Upload and process one or multiple PDF and ZDA files.
- Automatically identifies drilling equipment from serial numbers.
  - Known equipment aliases are mapped where configured.
  - New or unknown serial numbers are preserved as unique equipment instead of being grouped under a generic identifier.
- Reads cycle metadata, including:
  - cycle number
  - start date/time
  - serial number
  - drilling plan
  - section
  - drilled metres
- Extracts drilled-hole detail by type:
  - Bottom
  - Easer
  - Cut
  - Contour
  - Reaming
  - Casing
- Classifies drilling rounds using front-hole counts:
  - `FRENTE` > 45 holes
  - `SELLADA` 25–45 holes
  - `ESTOCADA Y/O CORRECCIONES` < 25 holes
- Uses front-hole types (`Bottom + Easer + Cut + Contour`) for round classification.
  - `Reaming` and `Casing` are not included in the front-hole count used for this classification.
- Analyzes drilling automation:
  - overall automatic movement by jumbo
  - automatic movement by boom
  - comparison between Boom 1 and Boom 2
- Analyzes drilled length in Cut holes:
  - cycle-level median
  - global median by jumbo
- Analyzes drilling-cycle timing from ZDA data:
  - real drilling start
  - real drilling end
  - drilling duration
  - average start time by jumbo and shift
- Reconstructs a reference drilling pattern from ZDA spatial data.
  - The reference contour is adjusted according to the detected drilling section.
  - The plot automatically expands when holes extend beyond the standard display area.
- Displays per-file and per-cycle validation details.
- Provides dynamic global filters by:
  - jumbo
  - drilling-round classification
- Exports a consolidated Excel workbook with:
  - `BD-PERFO`
  - `Resumen_Reportes`
  - `Resumen_Ciclos`

## Excel output

### `BD-PERFO`

Operational database by drilling cycle, including available fields such as:

- operational month
- date and shift
- operator
- jumbo
- level / block / heading
- drilling-round classification
- section
- rock class
- explosive type
- drilling and movement times
- automatic/manual movement indicators
- number of drilled holes
- cycle number
- automatic movement percentages
- Cut-hole median and average drilled length
- first-hole / first-hammer timing fields when available

The operational month uses a 26–25 reporting period. For example:

- August: July 26 to August 25
- September: August 26 to September 25

### `Resumen_Reportes`

One row per processed PDF or ZDA cycle, with cycle metadata, drilling totals, validation status, drilling-round classification, automation indicators, timing information, and reading-quality fields.

### `Resumen_Ciclos`

Cycle-level summary by hole type, including available statistics such as:

- count
- minimum
- maximum
- average
- median
- expected vs. extracted values
- validation status

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

## Validation

The application performs consistency checks between extracted drilling data and the information available in the source files.

If a cycle displays `REVISAR`, the extracted data should not be considered validated until the discrepancy is reviewed.

## Data handling

When deployed on Streamlit Community Cloud, uploaded PDF and ZDA files are processed on the Streamlit-hosted server.

Confirm that this is acceptable under your organization's information-security requirements before using operational, confidential, or sensitive files.
