
File Cleaner & Standardizer - README
====================================

Generated script: file_cleaner.py
Location: C:\Users\User\Documents\PruebaPiloto\scripts\file_cleaner.py

What it does:
    - A simple Tkinter GUI to scan folders, detect duplicates, validate and rename files, move and delete files.
    - Logs every action to actions_log.xlsx (if openpyxl is installed) or actions_log.csv as fallback.

How to run (Windows):
1. Install Python 3.8+ and optionally create a virtual environment.
2. (Optional but recommended) install openpyxl for Excel logging:
    pip install openpyxl
3. Run the script:
    python file_cleaner.py

Notes:
    - The GUI requires a desktop environment; it won't show inside a headless server session.
    - The script appends logs to actions_log.xlsx in the current working directory. If openpyxl is not installed, uses CSV fallback.
    - The rename template supports placeholders: {prefix}, {name}, {ext}, {date}, {counter}
    Example: {prefix}_{name}_{date}.{ext}

Security:
    - The script performs file operations (move/delete/rename). Always run on a copy or ensure backups exist before mass operations.
