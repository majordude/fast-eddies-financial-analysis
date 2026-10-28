FAST EDDIES FINANCIAL SPENDING AND INCOME ANALYSIS
READ THIS FIRST

What it does
  Turns bank and credit card statement PDFs into one Excel workbook: income vs expenses by month, spending by
  category and card, transfers, charts, and proof that every month ties to the statements. It also writes a short
  advisor summary (values only) that is safe to send.
  Two steps: (1) the redactor removes names, addresses and account numbers from your original PDFs; (2) the analyzer
  reads only the redacted copies. Nothing is uploaded anywhere by these scripts.


1. WHAT TO INSTALL

  Operating system
    Windows 11, 64-bit.

  PowerShell
    Nothing to install. Windows 11 includes Windows PowerShell 5.1, which runs every script here.
    Optional: PowerShell 7 (newer, not required). Install it with:  winget install -e --id Microsoft.PowerShell

  Python
    Python 3.12, 64-bit (recommended; 3.11 also works). Install with winget (built into Windows 11):
      winget install -e --id Python.Python.3.12
    Or download the installer from python.org/downloads and tick "Add python.exe to PATH" on the first screen.
    Close and reopen PowerShell afterwards, then check:  py --version

  Python libraries (setup_windows.ps1 installs all of these for you, listed in requirements.txt)
    pymupdf            reads the PDFs
    openpyxl           writes the Excel workbook
    presidio-analyzer  finds names and other personal data for redaction
    spacy              language engine used by presidio
    en_core_web_lg     spaCy English model, about 400 MB download (installed by setup_windows.ps1)

  Excel
    Microsoft Excel 2019 or newer, or LibreOffice Calc, to open the results.

  Git and GitHub CLI (optional)
    Only needed to download from, or publish to, GitHub. Without them, use GitHub's "Code > Download ZIP" button.
      winget install -e --id Git.Git
      winget install -e --id GitHub.cli
    Close and reopen PowerShell, then sign in once:  gh auth login


2. SET UP A NEW WINDOWS 11 PC

  a. Open PowerShell: press the Windows key, type PowerShell, press Enter.

  b. Create the folder and get the files. The standard location is C:\CCardStatements\Analysis.
       With git:
         mkdir C:\CCardStatements
         cd C:\CCardStatements
         gh repo clone <your-github-name>/<this-repo-name> Analysis
       Without git: download the ZIP from GitHub, extract it, and move the contents into C:\CCardStatements\Analysis.

  c. Run the one-time setup. It creates the folders, a private Python environment inside the Analysis folder, and
     installs the libraries and the spaCy model. Allow 10 to 20 minutes the first time (large download).
         cd C:\CCardStatements\Analysis
         powershell -ExecutionPolicy Bypass -File .\setup_windows.ps1

     "ExecutionPolicy Bypass" lets this one script run without changing any Windows setting. Windows blocks
     unsigned .ps1 files by default, which is why the command is written this way.

  d. (Optional) Bring your personal rules. personal_config.py holds your own vendors, payees, names and income
     sources. It is never stored in GitHub. Copy it by hand (USB drive, OneDrive, e-mail to yourself) into
     C:\CCardStatements\Analysis, next to analyze_statements.py. Also copy any of check_payees.json,
     vendor_overrides.json and brokerage_yearend.json that you use. Without personal_config.py the tool still
     works with generic categories; see personal_config.example.py to start one.


3. FOLDERS

  C:\CCardStatements\
      Inputs\     your original statement PDFs. Private. Never shared. Subfolders are fine (one per bank).
      Redacted\   redacted copies and any Excel/CSV spend lists. This is the only folder the analyzer reads.
      Analysis\   these scripts, the workbook, and processed_ledger.json (the memory of what was processed).

  To use a different location, set the environment variable FEFA_BASE before running, for example:
      $env:FEFA_BASE = "D:\Finance"


4. EVERY MONTH (short version)

  cd C:\CCardStatements\Analysis
  .\run_redact.ps1 -Test        (redacts the first PDF only: look at the result)
  .\run_redact.ps1              (redacts everything in Inputs)
  .\run_analyze.ps1             (builds the workbook in Analysis)

  If PowerShell refuses to run these, use:  powershell -ExecutionPolicy Bypass -File .\run_analyze.ps1
  Full steps, adding a new card, and troubleshooting are in HOWTO.txt. How the numbers are checked is in
  HOW_IT_WORKS.txt.


5. WHAT IS SUPPORTED

  Credit card statements: Citi (Costco Anywhere Visa, Premier layouts), Apple Card, and Synchrony store cards
  (Amazon Prime Store Card layout; each purchase line keeps its item name, shown on the Store Card Items tab).
  Bank statements: the credit-union layout this was built on (statement summary with member number, sub-accounts,
  date / description / amount / balance columns).
  Spend lists: any .xlsx or .csv with date, payee and amount columns.
  A different bank or card issuer needs its own reader; see "Adding a card" in HOWTO.txt.


6. IF SOMETHING FAILS

  "py is not recognized"            Python is not installed, or PowerShell was not reopened after installing it.
  "running scripts is disabled"     Use the "powershell -ExecutionPolicy Bypass -File ..." form shown above.
  pip or spaCy download errors      Check the internet connection and run setup_windows.ps1 again; it resumes.
  Excel says the file is locked     Close the workbook in Excel and run the analyzer again.
  Anything else                     The error is printed and saved to Analysis\last_error.txt.
