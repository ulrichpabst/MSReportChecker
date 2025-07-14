# MSReportChecker

**Real-Time HRMS Report Validation Directly Inside Microsoft Word**

MSReportChecker is a standalone Microsoft Word Add-in that automatically detects and evaluates High-Resolution Mass Spectrometry (HRMS) data within Word documents.

It validates formatting, extracts molecular formulas, calculates exact masses, and flags potential inconsistencies between reported and theoretical values — all in real time, without the need for copy-pasting or external software.

This tool is intended to support chemists, academic authors, and peer reviewers by ensuring HRMS data quality and adherence to reporting standards, particularly in manuscripts prepared for chemistry journals.


## Installation

The program is intended for use inside of Microsoft Word, compatibility is provided natively for both **macOS** and **Windows**. Since Linux is not natively supported by MS Word, the Add-in was not tested there. But – let's face it – if you manage to run MS Word on Linux, then you will most likely also succeed with this Add-in. ;-)


### macOS

Open the **Terminal**, then run:

```bash
# 1. Clone this repository
git clone https://github.com/ulrichpabst/MSReportChecker.git

# 2. Navigate into the folder
cd MSReportChecker

# 3. Run the installer
bash install_mac.sh
```
After installation, you may start the tool by running:
```bash
python server.py
```

### Windows

Open **PowerShell** as Administrator, then run:

```powershell
# 1. Clone the repository
git clone https://github.com/ulrichpabst/MSReportChecker.git

# 2. Navigate into the folder
cd MSReportChecker

# 3. Run the installer
.\install_win.ps1
```
After installation, you may start the tool by running:
```powershell
python server.py
```

Then in Microsoft Word:
1.	Go to `File > Options > Trust Center > Trust Center Settings`
2.	Choose `Trusted Add-in Catalogs`
3.	Add the full path to this project directory as a shared folder catalog
4.	Confirm and restart Word
5.	The MSReportChecker will now appear under `Home > Add-ins`


## Usage

1.  Ensure the local server is running:
```bash
python ./server.py
```
2.	Open Microsoft Word
3.	Select MS Report Checker from the Home ribbon tab
4.	Click Analyze Document in the taskpane to scan for HRMS reports
5.	Results are presented in a table with tooltips and colored status icons

## Security and Privacy

MSReportChecker is fully local and offline. It does not connect to the internet or transmit any data.
-	It runs a local server at `https://localhost:3000` to display the add-in interface.
-	A self-signed TLS certificate is created using [mkcert](https://github.com/FiloSottile/mkcert), a trusted open-source tool that installs a local certificate authority to enable secure HTTPS communication within your own computer only.
-	No telemetry, analytics, or user data is collected or stored.
-	You remain in full control of all documents and information.

## License

tba.

## Author

**Ulrich Pabst**, [ORCID](https://orcid.org/0009-0007-0529-0720), 2025
