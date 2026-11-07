# Programming for Data Science — Semester Assignment

Student number: 25106724  
Module: COM7024 Programming for Data Analytics

`com7024_analysis.py` runs the housing analysis. It imports the module dataset, checks types, removes rows that cannot be used as sales, and reports waterfront status, floor space, and year built against price.

## Run

Place `Manchester_house_Dataset[3678].csv` in this directory. The dataset is module material and is not stored in this repository.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python com7024_analysis.py
```

The script writes `results.json`, `execution_log.txt`, and the charts in `figures/`.
