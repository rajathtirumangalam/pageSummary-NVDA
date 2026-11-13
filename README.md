# Page Summary for NVDA

**A quicker way to understand what's inside a document.**

Page Summary is an NVDA add-on that helps screen reader users get a quick overview of the document they are currently reading. Instead of navigating through an entire page or document to understand its structure, users can request a summary and get useful information announced through NVDA.

The add-on supports web pages, PDF documents, e-books, and Microsoft Office applications.

## Features

### Web Pages
Get an overview of the current web page, including:
- Number of headings and their structure
- Landmarks
- Links
- Form fields
- Tables
- Lists
- Graphics

### PDF Documents
Get a structural overview of PDF documents, helping users understand the available content and document organisation.

### E-books
Get useful information about the current book page, including its structural elements.

### Microsoft Word
Get an overview of Word documents, including document structure and available content information.

### Microsoft Excel
Get information about Excel workbooks, including worksheets and workbook elements.

### Microsoft PowerPoint
Get an overview of presentations, including slide information and selected accessibility-related details.

## Keyboard Shortcuts

| Keyboard Shortcut | Action |
|---|---|
| NVDA + Shift + Q | Announce a summary of the current document |
| NVDA + Shift + Q (press twice quickly) | Open a detailed summary |
| NVDA + Ctrl + Shift + Q | Toggle automatic summaries |
| NVDA + Alt + Shift + Q | Display diagnostic information |

Keyboard gestures can be customised through:

**NVDA Menu → Preferences → Input Gestures → Page Summary**

## Automatic Summaries

Page Summary can automatically announce a document summary when a supported document loads.

Users can configure:
- Whether automatic announcements are enabled.
- Which document types should be announced.
- The delay before an announcement.
- Additional e-book reader applications.

## Settings

Open:

**NVDA Menu → Preferences → Settings → Page Summary**

The settings panel provides controls for automatic announcements, supported document types, announcement delay and additional e-book reader applications.

## Installation

1. Download the latest `.nvda-addon` package from the [Releases](https://github.com/rajathtirumangalam/pageSummary-NVDA/releases) section.
2. Open the downloaded add-on file.
3. NVDA will ask whether you want to install the add-on.
4. Confirm the installation.
5. Restart NVDA when prompted.

## Compatibility

- NVDA 2023.1 and later, subject to testing.
- Windows operating system.
- Microsoft desktop Office applications for Office-specific summaries.

Some features depend on the accessibility interfaces exposed by the application and document.

## Project Structure

```text
globalPlugins/
    pageSummary/
        __init__.py
        office.py
        report.py
        settings.py
        webdoc.py

doc/
    en/
        readme.html

manifest.ini
```

## Known Limitations

- Summary information depends on the document's underlying structure and accessibility information.
- PDF documents with poor tagging may provide limited structural information.
- Office-specific functionality requires compatible desktop Microsoft Office applications.
- Results may vary depending on the application and document format.

## Contributing

Contributions, bug reports, feature suggestions and accessibility testing feedback are welcome.

If you encounter an issue:
1. Check whether you are using the latest release.
2. Include your NVDA version.
3. Mention the application and document type.
4. Describe the steps required to reproduce the issue.
5. Include relevant NVDA log information where possible.

Please submit issues through the GitHub Issues section.

## Author

**Rajath Tirumangalam**

Accessibility Consultant | NVDA Add-on Developer

Website: https://rajathtirumangalam.com

GitHub: https://github.com/rajathtirumangalam

## License

This project is distributed under the GNU General Public License v2.0 (GPL-2.0).

See the [LICENSE](LICENSE) file for details.

---

*Page Summary – Spend less time figuring out a document. Spend more time understanding it.*