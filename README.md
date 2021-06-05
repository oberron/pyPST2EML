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

> python.exe -m pip install git+https://github.com/oberron/pyPST2EML


![][https://t4.ftcdn.net/jpg/01/01/48/09/240_F_101480925_dWEgCvJIagLTy36eiBmoyIRmxoqcKNeo.jpg]

once installed with pip you need to run the post install script

> python venv\Scripts\pywin32_postinstall.py -install

## USAGE

1. non regression testing on files in the `./test` folder

>python -m pyPST2EML -t Y

2. launch the conversion of .pst into a hiearical set of folders and .eml files

> python -m pyPST2EML --pst Y -f C:\outlook\archives\ -n 2021Q1.pst

3. Process an existing set of folders and .eml files to rename them / touch them

> python -m pyPST2EML --pst N -f C:\outlook\archives\eml\2021Q1


## NEXT STEPS

connect with a distributed database for fast querying of the database content. Example include but not limited to:

```
* rust
```
