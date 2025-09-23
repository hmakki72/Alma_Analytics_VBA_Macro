# Alma_Analytics_VBA_Macro
This VBA macro downloads and parses an XML report from Alma Analytics using a REST API. It dynamically extracts column headers and row data, then populates them into an Excel worksheet. Ideal for automating data retrieval and reporting in Excel or Power BI.
If you are new the Excel Developer, you may need to check these two links first:
 - Show the Developer tab: https://github.com/hmakki72/Alma_Analytics_VBA_Macro/tree/main
 - Enable or disable macros in Microsoft 365 files: https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-microsoft-365-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6

 
# Features 
- Connects to Alma Analytics via REST API
- Parses XML response using MSXML2.DOMDocument.6.0
- Dynamically extracts column headers from XML schema
- Populates data rows into Excel
- Handles errors gracefully (HTTP and XML parsing)

# Requirements
- Microsoft Excel (with VBA support)
- Alma Analytics API access
- Valid API key and report path
- Internet access

# Installation

- Open Excel.
- Press Alt + F11 to open the VBA Editor.
- Insert a new module.
- Paste the contents of ParseXMLWithDynamicHeaders into the module.
- Replace placeholders in the URL:
  - Your Analytics URL
  - Your report path
  - Your API key

# Usage
Run the macro ParseXMLWithDynamicHeaders, the macro will:
- Clear the first worksheet.
- Send a GET request to the Alma Analytics API.
- Parse the XML response.
Extract column headers and data rows.
Populate them into the worksheet.

# Notes
- The macro uses local-name() in XPath to handle XML namespaces.
- If saw-sql:columnHeading is missing, it falls back to the name attribute.
- Only the first worksheet (Sheets(1)) is used.

# Error Handling
Alerts if:
- HTTP request fails
- XML parsing fails
- No <Row> elements found

# License
- This project is open-source under the MIT License.

# Contributing
- Feel free to fork the repo, submit issues, or create pull requests to improve functionality or add features.

# Contact
For questions or support, please reach out via GitHub Issues or email.

** AI used to build this page

