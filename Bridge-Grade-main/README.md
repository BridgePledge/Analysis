# Bridge-Grade

## 3rd party dependencies

* requests
* bs4
* xlsxwriter

These can be installed either via pip et all, or by your OS (ie apt, yum).

## Usage

```
python3 scrape.py
```

## Organization

* `data`
    * `raw` - folder to hold scraped data
        * `commonground` - specifically HTML from commonground
    * `clean`
        * `bridge.sqlite3` - a DB file containing all loaded data sets.
        * `bridge.xlsx` - an  excel file containing all loaded data sets.
