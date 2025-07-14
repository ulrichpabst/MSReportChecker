# MSReportChecker

**Real-Time HRMS Report Validation Directly Inside Microsoft Word**

MSReportChecker is a standalone Microsoft Word Add-in that automatically detects and evaluates High-Resolution Mass Spectrometry (HRMS) data within Word documents according to the _ACS Guide to Scholarly Communication_:

<strong>HRMS</strong> (ESI/Q-TOF) m/z: [M + Na]<sup>+</sup> Calcd for C<sub>13</sub>H<sub>17</sub>NO<sub>3</sub>Na 258.1101; Found 258.1074.

It validates formatting, extracts molecular formulas, calculates exact masses, and flags potential inconsistencies between reported and theoretical values — all in real time, without the need for copy-pasting or external software.


This tool is intended to support chemists, academic authors, and peer reviewers by ensuring HRMS data quality and adherence to reporting standards, particularly in manuscripts prepared for chemistry journals.


## Installation

MSReportChecker is a Microsoft Word Add-in available for both **macOS** and **Windows**. It can be used in two ways:

- ✅ **Recommended**: Load the Add-in from GitHub Pages (auto-updating, zero setup)
- 💻 **Offline**: Use a local installation via a self-hosted server (fully local, no network access)

> 💡 Microsoft Word on Linux is not officially supported, but if you're managing to run it there, you probably know how to make the Add-in work too.

---
### ✅ Recommended Installation (GitHub Pages, cross-platform)

This is the easiest and preferred method.

#### macOS

1. Download the [MSReportChecker.xml](https://github.com/ulrichpabst/MSReportChecker/blob/main/MSReportChecker.xml) file  
2. Move it into the following directory:

   ```bash
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
3. Open Microsoft Word. The Add-in will appear automatically under the Home tab.

#### Windows

1. Download the [MSReportChecker.xml](https://github.com/ulrichpabst/MSReportChecker/blob/main/MSReportChecker.xml) file  
2. Open Microsoft Word, and go to `File > Options > Trust Center > Trust Center Settings`
3. Choose `Trusted Add-in Catalogs`
4.	Add the full path to the `MSReportChecker.xml` parent directory as a shared folder catalog in the `URL` box
5.	Confirm and restart Word
6.	The MSReportChecker will now appear under `Home > Add-ins`.

---
### 💻 Local Installation (Advanced/offline mode)

This is only needed for offline or air-gapped systems.

#### macOS

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

#### Windows

Open **PowerShell** as Administrator, then run:

```bash
# 1. Clone the repository
git clone https://github.com/ulrichpabst/MSReportChecker.git

# 2. Navigate into the folder
cd MSReportChecker

# 3. Run the installer
.\install_win.ps1
```
After installation, you may start the tool by running:
```bash
python server.py
```

Then in Microsoft Word:
1.	Go to `File > Options > Trust Center > Trust Center Settings`
2.	Choose `Trusted Add-in Catalogs`
3.	Add the full path to this project directory as a shared folder catalog
4.	Confirm and restart Word
5.	The MSReportChecker will now appear under `Home > Add-ins`



## Usage

### ✅ Default (GitHub Pages)

1. Open **Microsoft Word**.
2. Ensure the Add-in is installed and visible under the **Home** tab.
3. Click **MS Report Checker** to open the taskpane.
4. Click **Analyze Document** to scan the open file for HRMS report lines.
5. The results are presented in a structured table, with tooltips and color-coded status icons to highlight syntax issues, mass errors, or formatting deviations.

---

### 💻 Offline Mode (Localhost version)

> Only applicable if you chose the advanced local installation method.

1. Start the local server:
   ```bash
   ./server.py

2.	Open Microsoft Word.
3.	Click the MS Report Checker button in the ribbon.
4.	Use the taskpane as described above.


## Security and Privacy

MSReportChecker is designed to operate with **maximum data integrity and privacy**. It does not collect, transmit, or expose any document content, regardless of installation method.

---
### ✅ GitHub Pages Version (Recommended)

- The add-in interface is loaded securely over **HTTPS** from GitHub Pages (`https://ulrichpabst.github.io/MSReportChecker/`).
- **No document content is ever uploaded or transmitted** to any server.
- All analysis is performed entirely within your local Word environment, using the [Office JavaScript API](https://learn.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office).
- No telemetry, tracking, cookies, or analytics are used.
- All source code is publicly auditable, and you retain full control over the manifest and access paths.

---
### Local Installation (Offline Mode)

- The add-in is hosted locally via `https://localhost:3000`, using a self-signed TLS certificate generated by [mkcert](https://github.com/FiloSottile/mkcert).
- This mode is suitable for **air-gapped systems** or environments where internet access is restricted or forbidden.
- The interface is served from your own machine; no network requests are made.
- All functionality remains fully offline and under your control.

---

In both configurations:

- MSReportChecker does **not transmit** any user data, document content, formulas, or metadata.
- You may inspect the source code to verify that all processing occurs client-side.
- The add-in requests only the `ReadWriteDocument` permission, necessary to read text and insert taskpane content. It does **not** access files, external storage, or other system resources.


## License

tba.

## Author

**Ulrich Pabst**, [ORCID](https://orcid.org/0009-0007-0529-0720), 2025
