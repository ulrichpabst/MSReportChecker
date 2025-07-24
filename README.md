# MSReportChecker

**Real-Time HRMS Report Validation Directly Inside Microsoft Word**

MSReportChecker is a standalone Microsoft Word Add-in that automatically detects and evaluates High-Resolution Mass Spectrometry (HRMS) data within Word documents according to the _ACS Guide to Scholarly Communication_:

<strong>HRMS</strong> (ESI/Q-TOF) m/z: [M + Na]<sup>+</sup> Calcd for C<sub>13</sub>H<sub>17</sub>NO<sub>3</sub>Na 258.1101; Found 258.1074.

It validates formatting, extracts molecular formulas, calculates exact masses, and flags potential inconsistencies between reported and theoretical values — all in real time, without the need for copy-pasting or external software.


This tool is intended to support chemists, academic authors, and peer reviewers by ensuring HRMS data quality and adherence to reporting standards, particularly in manuscripts prepared for chemistry journals.


## Installation

MSReportChecker is a Microsoft Word Add-in available for both **macOS** and **Windows** with stable support starting from MS Office 2021. It can be used in two ways:

- ✅ **Recommended**: Load the Add-in from GitHub Pages (auto-updating, zero setup)
- 💻 **Offline**: Use a local installation via a self-hosted server (fully local, no network access)

Microsoft Word on Linux is not officially supported, but if you're managing to run it there, you probably know how to make the Add-in work too.

Here, the installation of the recommended form is presented. To install the offline version, please clone this repository and run the installer scripts for your specific platform.

### macOS

1. Download the [`MSReportChecker.xml`](https://raw.githubusercontent.com/ulrichpabst/MSReportChecker/main/MSReportChecker.xml) manifest file.
2. Move it into the following directory:

   ```bash
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
3. Open Microsoft Word. The Add-in will now appear under `Insert > Add-Ins`. The precise location of the Add-in might differ between different versions of Word.

### Windows

1. Download the [`MSReportChecker.xml`](https://raw.githubusercontent.com/ulrichpabst/MSReportChecker/main/MSReportChecker.xml) manifest file.
2. Move the file into any directory.
3. Open the `Properties` Settings of the directory, move to the `Network` tab, then grant access to the specific user/group (requires the Administrator password).
   > If you do not own your computer and all rights (e.g., an institutional device), your responsible IT service may help you move the XML file to something like a shared directory and obtain a respective valid network access URL. This may very well be the trickiest part of the whole installation process.    
5. Copy the path for the shared location of the directory (e.g., file://Path/To/Directory)
6. Open Microsoft Word, and go to `File > Options > Trust Center > Trust Center Settings`
7. Choose `Trusted Add-in Catalogs`
8.	Add the full path to this project directory as a shared folder catalog in the `URL` field, but make sure to remove trailing `)` characters and everything that comes before `//`. Check the Option `Show in Menu`.
9.	Confirm and restart Word.
10.	The MSReportChecker will now appear under `Insert > My Add-ins > Shared Folder`. The precise location of the Add-in might differ between different versions of Word.









## Usage

1.	Open a Microsoft Word Document
2.	Select MS Report Checker from the Home ribbon tab
3.	Click `Analyze Document` in the taskpane to scan for HRMS reports
4.	Results are presented in a table with tooltips and colored status icons.

You can jump to the location of specific reports in the document by clicking on the linked number (left in each entry row).

Furthermore, if you hover above the respective icon, contextual advice will be presented for correcting the problematic entry.

## Security and Privacy

- The add-in interface is loaded securely over **HTTPS** from GitHub Pages (`https://ulrichpabst.github.io/MSReportChecker/`).
- **No document content is ever uploaded or transmitted** to any server.
- All analysis is performed entirely within your local Word environment, using the [Office JavaScript API](https://learn.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office).
- No telemetry, tracking, cookies, or analytics are used.
- All source code is publicly auditable at all times, and you retain full control over the manifest and access paths.


## License

tba.

## Author

[**Ulrich Pabst**](https://orcid.org/0009-0007-0529-0720), 2025
