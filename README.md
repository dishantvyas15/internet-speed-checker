# internet-speed-checker

This Python script tests the speed of your internet connection and records the value in an Excel sheet.

Please note that the Excel sheet present in this repo must be used and placed in the same folder as the Python script to avoid formatting clashes while storing test results.<br>

The sample execution log text file is only for reference. The actual execution log file will be created during the first run of the script after which, it'll be updated on every execution.

Requirements:<br>
--> Python<br>
--> `speedtest-cli` package<br>
--> `openpyxl` package

<hr>
Python Installation:<br>
Download and install the latest Python version from https://www.python.org/downloads/<br>
Tutorial: https://youtu.be/xXEt9dyvq3U<br>
Make sure you check the options of installing <i>pip</i> (which would be required to install other Python packages) and <i>Add Python to environment variables</i>.<br><br>

You may check your Python installation by running `python --version` in terminal.<br>
The Python packages can be installed by running `pip install <package-name>` in terminal.<br>

<hr>
This script may be automated using <i>Task Scheduler</i> (in Windows) in order to get test results periodically. I found this tutorial useful for the same: https://youtu.be/DVUlkU2AxgQ<br>
Use the <b>Create Task...</b> option instead of <b>Create Basic Task...</b> option to get more customization options.
