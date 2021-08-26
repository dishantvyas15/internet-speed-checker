# internet-speed-checker

This Python script tests the speed of your internet connection and records the value in an Excel sheet.

Please note that the Excel sheet present in this repo must be used and placed in the same folder as the Python script to avoid formatting clashes while storing test results.<br>

The sample execution log text file is only for reference. The actual execution log file will be created during the first run of the script after which, it'll be updated on every execution.

Requirements:<br>
--> Python<br>
--> `speedtest-cli` package<br>
--> `openpyxl` package

The Python packages can be istalled by running the following command:<br>
`pip install <package-name>`

This script may be automated using Windows Task Manager in order to get test results for multiple instances.
