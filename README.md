# COM7024 Manchester housing analysis

Student number: 25106724

Python analysis for the Programming for Data Analytics project. The script profiles the module housing file, cleans rows that cannot be used as sales, and tests waterfront status, floor space, and year built against price.

## Run

Place `Manchester_house_Dataset[3678].csv` in this directory. That file is module material and is not included in this repository.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python com7024_analysis.py
```

Outputs are written beside the script: `results.json`, `execution_log.txt`, and `figures/`.

`build_report.py` rebuilds the Word submission. It needs the official Arden cover page and the outputs of the analysis script, so it is meant to be run locally rather than as a public demo.
