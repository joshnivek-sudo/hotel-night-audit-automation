# Hotel Night Audit Automation

![Python](https://img.shields.io/badge/Python-3-3776AB?logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-217346)
![Tests](https://img.shields.io/badge/tests-28%20passing-brightgreen)
![License MIT](https://img.shields.io/badge/license-MIT-blue)

Python scripts that automate two daily routines of a hotel's night audit: the **card payment reconciliation** and the **daily statistics report**. Reports exported from the hotel management system (PMS) go in; finished spreadsheets come out, ready to send.

🇧🇷 *[Leia em português](README.md)*

---

## Results

| Routine | Before | After |
|---|---|---|
| 💳 Card reconciliation | 20–25 min | **< 2 min** |
| 📊 Statistics + management summary | 30–40 min | **< 2 min** |
| **Per night** | **almost 1 hour** | **under 4 minutes** |

With fewer errors, too: what used to rely on double-checking by eye now has automatic cross-checks.

## Context

I work as a night auditor at a hotel in Brasília, Brazil. Every night, part of the shift went into copying numbers from system reports into Excel: totaling card payments by brand, separating card-terminal payments from online ones, switching between four reports to fill in the statistics — and then reviewing everything, because one misplaced number ruins the close.

Repetitive work, well-defined rules and plenty of room for typos: a textbook case for automation. I built these tools for my own shift, and the card reconciliation is now also used by another auditor on the team.

## The two automations

### 💳 Card Payment Reconciliation
Reads the daily transactions report (`.xlsx` or `.pdf`), identifies each entry's card brand, separates **card terminal** from **online payment** and fills in the reconciliation sheet. Formulas compute each brand's total and online share; the auditor only types in the terminal closings.

→ [Business rules](docs/cartoes.md) *(Portuguese)*

### 📊 Daily Statistics Report
Combines four PMS reports (daily summary, on-the-books forecast, pending accounts and live room status) to fill in the statistics workbook and the one-page management summary — occupancy, ADR, RevPAR, revenue, 3-day forecast, pending accounts, walk-ins, no-shows. It then checks that the revenue it filled in matches the system's total.

→ [Field mapping](docs/estatistico.md) *(Portuguese)*

## Try it yourself

The repository ships with **100% fictitious reports and spreadsheets** in [`exemplos/`](exemplos/), in the same format the scripts read.

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v    # 28 tests
```

Step-by-step demo instructions are in [`exemplos/README.md`](exemplos/README.md) *(Portuguese)*.

## Tech stack

- **Python 3**, no frameworks
- **openpyxl** — filling in the existing Excel workbooks
- **xlrd** — reading legacy `.xls` reports
- **pypdf** — reading the transactions report as PDF
- **zipfile + xml.etree** — hand-parsing an `.xlsx` that openpyxl refuses to open
- **unittest** — automated tests
- **Windows `.bat`** — double-click to run, no terminal needed

## Technical challenges

- **An `.xlsx` that Excel opens but Python won't.** The PMS export has a misspelled XML attribute and openpyxl rejects it. The fix was to treat the file as what it is — a zip of XML files — and read the data directly.
- **Messy source data.** Undescribed entries, recurring typos (`B2PAY` for `BEE2PAY`) and acquirer names glued to card brands (`"CIELO"` contains `"ELO"`). Each case became a rule — and a test.
- **A rule that didn't hold up.** The first version identified card terminals by the document number. Over time the pattern changed, so the rule was replaced by one based only on what staff actually write.
- **Finding values by label, not position.** Reports are read as a grid and each field is located by its label, so the script keeps working if a row moves.
- **Two sources, two points in time.** The daily summary reflects the close; room status is a live snapshot. Each field uses the source that matches what it represents.
- **Filtering out what isn't the hotel's.** The building shares space with condo units that appear in the same report and must be excluded.
- **Distrusting its own output.** Revenue is reconciled against the system total, mismatches raise warnings, files are backed up before saving and written atomically.
- **Non-technical users.** Everything runs with a double click. The statistics tool is a single file that is both a Windows batch script and a valid Python program.

## Data privacy

Real reports contain guest names and financial data, and the hotel's spreadsheets belong to the hotel. So there is **no real data** in this repository, the examples are **invented**, the spreadsheet templates were **designed from scratch** for the demo, and the code contains no hotel names or terminal numbers.

## Author

**Kevin Joshua Siqueira Marques** — transitioning into data analysis, Systems Analysis and Development student.

[LinkedIn](https://br.linkedin.com/in/kevin-joshua-291b26249)
