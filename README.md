#TrackerSleep

The program is designed for sleep schedule tracking and takes the following parameters into account:

1. Bedtime (date day:month:year Exact time).
2. Wake-up time (date day:month:year Exact time).
3. Sleep duration.
4. Wake duration.
5. Shift — how much later or earlier bedtime was compared to the previous day.
6. Day — the actual length of the sleep-wake cycle (subjective day).
7. Comment.

The program was created for sleep accounting in cases of chronic sleep schedule disruption, for medical and/or research tracking.
1. It builds an Excel spreadsheet and draws a chart/graph.
2. It tracks the following average parameters:

Average sleep duration, wake duration, bedtime shift, and day length.

Minimum and maximum sleep duration, wake duration, bedtime shift, and day length.

The main version is designed for internet geeks; it automatically counts any time outside Google Chrome lasting longer than 3 hours as sleep.
Based on pauses in browser history and Windows program logs, it automatically calculates sleep periods and presents them to the user (incorrect entries can be deleted).
Calculation starts from the last date in the Excel file.

The second version (marked with the number 1) allows manual date/time entry and adding comments.

For installation, create a folder, download the desired program version file, the .bat file (which launches the program with a double click), and requirements.txt.

1. Open Windows PowerShell (Python must be installed).

2. Navigate to the folder: cd desired_folder

3. Create a virtual environment: python -m venv .venv

4. Activate the environment: ./.venv/Scripts/Activate.ps1

5. Install dependencies: pip install -r requirements.txt

The program is only available in Russian.

Developed on Windows 11; compatibility with other operating systems has not been tested.
