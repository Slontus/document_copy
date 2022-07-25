# Script to perform snips from protected DRM file

This is script which allows PDF pages turning and screen capturing with saving to Word document.

### How to Install

Python3 should be already installed. Then use pip (or pip3,
if there is a conflict with Python2) to install dependencies:
```
pip install -r requirements.txt
```
To run this program you need to find your PDF viewer process ID first:
on Windows run the following command in PowerShell:

Get-Process

find your application process ID and copy it in the script.ps1 file