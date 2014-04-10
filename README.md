# Netfile ETL on Windows

Replicate the ETL process from Dave Guarino on a Windows environment. Process .zip files from the [Netfile](https://www.netfile.com/) campaign finance filing system into individual CSVs per form.

## What It Does

Python script to perform the following:

1. Downloads and unzips the .zip files

2. Extracts the individual sheets from each Excel workbook as standalone CSVs

3. Combines the sheets for each "form" (for example, "A-Contributions") across all years, yielding a single CSV with all years' data for each individual form.


## Getting Started
Running the python scripts in a Windows system requires installing additional applications. 64 bit versions were installed when available on the current server , example: curl, python.

### Installing Dependencies
For running on Windows the following dependencies include options for compiling from source or downloading precomiled versions to install.

#### System Dependencies
- [curl](http://curl.haxx.se/)
- [unzip](http://gnuwin32.sourceforge.net/packages/unzip.htm)
- [gawk](http://gnuwin32.sourceforge.net/packages/gawk.htm)
- [python 2.7](http://www.python.org)


#### Python Dependencies
- [pip](http://www.pip-installer.org)
- csvkit

### Installation
Install curl, unzip, and gawk from precompiled installers avaiable online.

Install python 2.7
Set system path for python, added using powershell:
[Environment]::SetEnvironmentVariable("Path", "$env:Path;C:\Python27\;C:\Python27\Scripts\", "User")

Install PIP by downloading [get-pip.py](http://www.pip-installer.org/en/latest/installing.html) and running from the command prompt "python get-pip.py". 
Sample run file from python folder:
- Open Command Prompt
- cd\python27
- python get-pip.py

Install csvkit using PIP
- Open Command Prompt
- pip install csvkit

Note: initial install for csvkit returned error:
Could not find a version that satisfies the requirement argparse==1.2.1 (from -r requirements.txt (line 4)) (from versions: 0.1.0, 0.2.0, 0.3.0, 0.4.0, 0.5.0, 0.6.0, 0.7.0, 0.8.0, 0.9.0, 0.9.1, 1.0.1, 1.0, 1.1)
Some externally hosted files were ignored (use --allow-external to allow).

Clear error by installing latest version:
pip install argparse --allow-external argparse --upgrade

### Running

Task Sceduler is used to run the script on a scheule.
