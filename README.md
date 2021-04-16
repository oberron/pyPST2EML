# README
This small package leverages libpst to turn PST files into directories of .eml and .ics files which can then be indexed by desktop search tools.

The .eml and .ics files have their properties (author, creation date, and file name changed to reflect the email properties for easier search queries).

## REQUIREMENTS
* LIBPST binaries for windows (can be found on the internet)
* python (tested with python 3.8 32b on Win10)

## INSTALL

```bash
python -m venv venv_eml
.\venv_eml\Scripts\activate
pip install -r requirements.txt
```

or 
```
python.exe -m pip install git+https://github.com/oberron/pyPST2EML
```

## USAGE
