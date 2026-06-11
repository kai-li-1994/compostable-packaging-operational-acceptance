# Certified but not accepted: operational limits to compostable packaging in organic-waste systems

This repository contains the archived evidence corpus, Python coding workflow, figure-generation scripts, and Supplementary Data generation workflow for the manuscript:

**Certified but not accepted: operational limits to compostable packaging in organic-waste systems**

**Release and archive**

- GitHub repository: https://github.com/kai-li-1994/compostable-packaging-operational-acceptance
- Zenodo archive: https://doi.org/10.5281/zenodo.20650603

The study examines whether compostable packaging that is marketed or certified as compostable is actually accepted by organic-waste collection and treatment systems. The workflow converts a corpus of public operational sources into source-level coding outputs, source × application-group records, certification-sufficiency outputs, figure files, and a reader-facing Supplementary Data workbook.

## Authors and project context

This repository was prepared by **Kai Li** and **Grit Walther**.

- **Kai Li** is a postdoctoral researcher at the Chair of Operations Management, RWTH Aachen University, Germany, and is affiliated with the Institute of Environmental Sciences (CML), Leiden University, the Netherlands.
- **Grit Walther** is Professor and Chair of Operations Management at RWTH Aachen University, Germany.

The repository supports reproducible analysis for a study on the operational governance of compostable packaging in organic-waste systems. The analysis focuses on public operational rules, including municipal sorting instructions, waste-operator guidance, treatment-facility rules, and national or regional policy documents.

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
```

A minimal environment requires:

```text
python >= 3.10
pandas
numpy
matplotlib
scipy
openpyxl
```

The scripts do not require internet access after `web_sources.zip` is available locally.

## Reproduce the analysis

Run the full workflow from the repository root.

```bash
python result_calculation.py --input web_sources.zip --outdir results
python supplementary_data_generation.py --input results --output Supplementary_Data.xlsx
python figure_1_schematic.py
python figure_2_source_level_operational_rules.py
python figure_3_application_compatibility_country_profiles.py
python figure_4_certification_sufficiency.py
```

## Script details

### `result_calculation.py`

This is the core analytical script. It reads the archived source corpus and generates standardized coding outputs.

Default command:

```bash
python result_calculation.py --input web_sources.zip --outdir results
```

Main outputs:

| Output folder | Description |
|---|---|
| `results/display/` | Reader-facing CSV tables used by the figures and Supplementary Data workbook. |
| `results/audit/` | Long-form regex/rule traceability outputs, matched-rule information, and validation checks. |

### `supplementary_data_generation.py`

This script is the presentation layer for the Supplementary Data workbook. It reads the display CSV files generated by `result_calculation.py`, selects reader-facing sheets, masks local paths, removes internal matched-text trace columns from clean sheets, and formats the workbook. It does not change the analytical results.

Default command:

```bash
python supplementary_data_generation.py --input results --output Supplementary_Data.xlsx
```

### Figure scripts

Each figure script reads from `results/display/` and writes outputs to `figures/`.

| Script | Main outputs |
|---|---|
| `figure_1_schematic.py` | `figures/figure_1_schematic.pdf`, `.svg`, `.png` |
| `figure_2_source_level_operational_rules.py` | `figures/figure_2_source_level_operational_rules.pdf`, `.png`, `.csv` |
| `figure_3_application_compatibility_country_profiles.py` | `figures/figure_3_application_compatibility_country_profiles.pdf`, `.png`, status/order CSVs |
| `figure_4_certification_sufficiency.py` | `figures/figure_4_certification_sufficiency.pdf`, `.png`, `.csv` |

The PDF/SVG figure outputs are intended for vector editing and journal production where applicable.

## Analytical outputs

The workflow produces outputs for three analytical routes.

| Analytical route | Main question | Main output types |
|---|---|---|
| Source-level operational rules | Who defines disposal rules for compostable packaging? | Source authority, stated treatment route, acceptance outcome, rejection rationale. |
| Application-level compatibility | Which compostable-packaging applications are accepted, restricted, or rejected? | Source × application-group records and lower, central, upper compatibility scenarios. |
| Certification sufficiency | Is named certification or approval sufficient to secure disposal acceptance? | Certification-relevant records by application group, certification basis, source authority, and sufficiency status. |

## Supplementary Data workbook

`Supplementary_Data.xlsx` is the reader-facing supplementary data file associated with the manuscript. It includes:

- result summary sheets,
- source-level coding,
- source × application-group records,
- application compatibility scenario outputs,
- certification-sufficiency outputs,
- country coverage information,
- sensitivity checks,
- codebook sheets with category definitions and representative regex examples.

## Data and source-archive note

The derived coding outputs, Supplementary Data workbook, and Python scripts are shared for transparency and reproducibility. The archived source files in `web_sources.zip` are copies of public third-party webpages and PDFs retained to support verification of the analysis. Rights in those source materials remain with their original publishers. Users should consult the original source publishers for reuse beyond verification, reproducibility, or scholarly review.

## License

The Python code in this repository is released under the MIT License, provided in `LICENSE`.

The derived coding outputs and `Supplementary_Data.xlsx` are shared to support transparency, verification, and reuse of the analysis. The archived public source files in `web_sources.zip` are not relicensed by this repository. They are copies of public third-party webpages and PDFs retained as an evidence corpus for verification and reproducibility, and rights in those source materials remain with their original publishers.

## Citation

If you use this repository or the archived release, please cite the GitHub repository and the Zenodo archive.

```text
Li, K., & Walther, G. (2026). Certified but not accepted: operational limits to compostable packaging in organic-waste systems: reproducible data and code archive (v1.0.0). Zenodo. https://doi.org/10.5281/zenodo.20650603
```

Zenodo archive: https://doi.org/10.5281/zenodo.20650603  
GitHub repository: https://github.com/kai-li-1994/compostable-packaging-operational-acceptance

## Acknowledgements

This repository supports research conducted at the Chair of Operations Management, RWTH Aachen University, with affiliation to the Institute of Environmental Sciences (CML), Leiden University. The authors thank the public authorities, waste operators, treatment facilities, and regional or national organisations whose publicly available operational guidance made the evidence corpus possible.

This research was supported by the Werner Siemens Foundation through the WSS Research Centre Catalaix, a Project of the Century funded by the Werner Siemens Foundation.

## Contact

For questions about the repository or manuscript, contact:

**Kai Li**  
Chair of Operations Management, RWTH Aachen University  
Institute of Environmental Sciences (CML), Leiden University  
Email: `kai.li@om.rwth-aachen.de`