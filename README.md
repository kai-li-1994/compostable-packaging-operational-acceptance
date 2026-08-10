[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.20650603.svg)](https://doi.org/10.5281/zenodo.20650603)

# Certified but not accepted: operational limits to compostable packaging in organic-waste systems

This repository contains the archived evidence corpus, Python coding workflow, figure-generation scripts, generated analytical outputs, and Supplementary Data generation workflow for the manuscript:

**Certified but not accepted: operational limits to compostable packaging in organic-waste systems**

## Release and archive

- GitHub repository: https://github.com/kai-li-1994/compostable-packaging-operational-acceptance
- Zenodo archive: https://doi.org/10.5281/zenodo.20650603

The study examines whether compostable packaging that is marketed or certified as compostable is actually accepted by organic-waste collection and treatment systems. The workflow converts a corpus of public operational sources into source-level coding outputs, source × application-group records, certification-sufficiency outputs, figure files, and a reader-facing Supplementary Data workbook.

## Authors and project context

This repository was prepared by **Kai Li** and **Grit Walther**.

- **Kai Li** is a postdoctoral researcher at the Chair of Operations Management, RWTH Aachen University, Germany, and is affiliated with the Institute of Environmental Sciences (CML), Leiden University, the Netherlands.
- **Grit Walther** is Professor and Chair of Operations Management at RWTH Aachen University, Germany.

The repository supports reproducible analysis for a study on the operational governance of compostable packaging in organic-waste systems. The analysis focuses on public operational rules, including municipal sorting instructions, waste-operator guidance, treatment-facility rules, regional rules, and national policy documents.

## Repository structure

| Path | Type | Role |
|---|---|---|
| `web_sources.zip` | Input archive | Original archived evidence corpus used as the stable input to the coding workflow. |
| `result_calculation.py` | Core script | Reads `web_sources.zip`, applies rule-based coding, and generates source-level, application-level, certification-sufficiency, sensitivity, audit, and codebook CSV outputs. |
| `supplementary_data_generation.py` | Core script | Converts the generated display CSV files into the reader-facing `Supplementary_Data.xlsx` workbook. |
| `figure_1_schematic.py` | Figure script | Generates Figure 1, the analytical workflow schematic. |
| `figure_2_source_level_operational_rules.py` | Figure script | Generates Figure 2, source-level operational rules. |
| `figure_3_application_compatibility_country_profiles.py` | Figure script | Generates Figure 3, application-level compatibility and country profiles. |
| `figure_4_certification_sufficiency.py` | Figure script | Generates Figure 4, certification and certification-sufficiency results. |
| `Supplementary_Data.xlsx` | Output workbook | Reader-facing supplementary data file accompanying the manuscript. It can be regenerated from `results/display/`. |
| `results/display/` | Generated outputs | Reader-facing CSV outputs used by figures and Supplementary Data. Created by `result_calculation.py`. |
| `results/audit/` | Generated outputs | Regex/rule traceability files and validation/audit outputs. Created by `result_calculation.py`. |
| `figures/` | Generated outputs | PNG/PDF/SVG figure files and figure-data CSVs. Created by the figure scripts. |
| `requirements.txt` | Environment file | Minimal Python package requirements. |
| `LICENSE` | License file | Code license. See also the data and source-archive note below. |

The `results/` and `figures/` folders are generated outputs. They can be deleted and recreated by running the scripts described below. The final `Supplementary_Data.xlsx` file is included for convenience and can also be regenerated.

## Evidence archive: `web_sources.zip`

The file `web_sources.zip` is the archived source corpus used by the reproducible workflow. It contains one root folder, `web_sources_log_free/`, with the following structure.

| Path inside `web_sources.zip` | Content | Role in workflow |
|---|---|---|
| `web_sources_log_free/01_references_corpus.jsonl` | Consolidated JSON Lines corpus | Stable input read by `result_calculation.py`. Each line represents one retained public operational or policy source. |
| `web_sources_log_free/archive_manifest.csv` | Archive manifest | Lists archived source files and supports traceability of the corpus. |
| `web_sources_log_free/metadata/` | Per-source JSON metadata files | Records source-level attributes, source identifiers, archive/access information, and related metadata. |
| `web_sources_log_free/raw_html/` | Archived HTML files | Raw archived webpages where available. |
| `web_sources_log_free/raw_pdf/` | Archived PDF files | Raw archived PDFs where available. |
| `web_sources_log_free/text/` | Extracted text files | Text extracted from archived sources and used for regex-based coding. |

The corpus preserves operational source wording, including local-language terminology for organic-waste streams, accepted and rejected items, certification labels, treatment routes, and contamination warnings.

## Workflow overview

The workflow has three main stages.

1. **Result calculation**  
   `result_calculation.py` reads `web_sources.zip`, applies indicator-specific rule dictionaries, and writes display and audit outputs to `results/`.

2. **Supplementary Data generation**  
   `supplementary_data_generation.py` reads the generated display CSV files and creates `Supplementary_Data.xlsx`.

3. **Figure generation**  
   The four figure scripts read the display CSV files and generate the manuscript figures in `figures/`.

## Installation

Create a clean Python environment and install the required packages.

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows PowerShell
pip install -r requirements.txt
