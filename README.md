# AutoDoc Genius

**English** | [简体中文](./README.zh-CN.md)

## Project Overview

**AutoDoc Genius** is a desktop application built with Electron + Vue 3 for batch generating Word/PDF reports. The tool reads multiple rows of data from an Excel spreadsheet and combines them with a Word template file to automatically generate the corresponding number of Word or PDF report files.

### Core Features

1. **Excel Data Reading**
   - Supports .xlsx and .xls formats
   - Automatically parses headers and data rows
   - Automatic handling of date and time formats

2. **Word Template Processing**
   - Supports .docx template files
   - Placeholder replacement mechanism (using markers such as XXX01, XXX02)
   - Dynamically generates personalized documents

3. **Multi-Option Support**
   - Supports checkbox template functionality
   - Uses `{{check01XXX}}{{check02XXX}}` style markers for multi-option placeholders
   - In Excel, separate multiple options with commas, Chinese commas, or the enumeration comma (、)
   - Automatically marks selected checkboxes with `☑` and keeps unselected ones as `□`

4. **Image Replacement Support**
   - Excel columns whose headers start with `img` (case-insensitive) are treated as image columns (e.g., `img01`)
   - Cells contain image file paths; both absolute paths and paths relative to the Excel file's directory are supported
   - Image placeholders (e.g., `img01`) are written directly in the Word template and replaced with real images during generation
   - Supports png / jpg / jpeg / gif / bmp; images are embedded at their original aspect ratio and scaled down proportionally when wider than 15cm
   - Missing images leave the position empty and produce a warning

5. **Batch Document Generation**
   - Supports export to Word (.docx) format
   - Supports export to PDF format (requires an Office environment)
   - Progress bar shows generation progress
   - Automatically generates a ZIP archive

6. **Office Integration**
   - Detects Microsoft Word installation status
   - Supports WPS Office as an alternative
   - Invokes the COM interface via PowerShell scripts

## Project Structure

```
vite-project/
├── electron/                    # Electron main process code
│   ├── main.js                 # Main process entry file
│   ├── preload.js              # Preload script
│   └── officeConverter.js      # Office document conversion module
├── src/
│   ├── utils/                  # Utility functions
│   │   ├── documentUtils.js    # Document processing utilities
│   │   ├── docxGenerator.js    # Word document generator
│   │   ├── imageUtils.js       # Image replacement utilities
│   │   └── fileUtils.js        # File operation utilities
│   ├── App.vue                 # Main application component
│   └── main.js                 # Vue application entry
├── public/                     # Static assets
│   ├── Excel Template.xls      # Sample Excel template
│   ├── Word Template.docx      # Sample Word template
│   └── vite.svg                # Vite icon
├── .vscode/                    # VS Code configuration
├── package.json                # Project configuration
├── vite.config.js              # Vite configuration file
├── auto-imports.d.ts           # Auto-import type definitions
└── components.d.ts             # Component type definitions
```

## Installation & Running

### Install Dependencies

```bash
npm install
```

### Development Mode

```bash
npm run dev
```

### Build the Application

```bash
npm run build
```

After the build completes, the installer package will be generated in the `release` directory.

## Usage

1. **Prepare the Excel Spreadsheet**
   - Row 1: Description information (optional)
   - Row 2: Field names (used for template placeholders)
   - Row 3 onwards: Data rows

2. **Prepare the Word Template**
   - Use markers such as XXX01, XXX02 for regular placeholders
   - Placeholders must correspond to the Excel field names
   - Use the `{{checkXXX}}` format to mark multi-option placeholders
   - Inside `{{checkXXX}}`, define options using the `□ Option1 □ Option2 □ Option3` format

3. **Multi-Option Usage Example**
   - **Excel header**: `check01`
   - **Excel data**: `Swimming,Running,Reading` (separated by commas, Chinese commas, or the enumeration comma)
   - **Word template**: `{{check01□ Swimming □ Running □ Reading □ Programming}}`
   - **Generated result**: `☑ Swimming ☑ Running ☑ Reading □ Programming`

4. **Image Replacement Usage Example**
   - **Excel header**: `img01`
   - **Excel data**: `photos/photo1.jpg` (relative to the Excel file's directory) or `D:\pics\photo1.jpg` (absolute path)
   - **Word template**: write `img01` directly where the image should appear (preferably in its own paragraph)
   - **Generated result**: the placeholder is replaced with the referenced image

5. **Upload Files and Generate**
   - Upload the Excel file and the Word template
   - Select the export format (Word or PDF)
   - Click the generate button
   - Download the generated ZIP file

## Notes

- The PDF export feature requires Microsoft Word or WPS Office to be installed
- Make sure the placeholders in the template exactly match the Excel headers
- Generating a large batch of documents may take a while, please be patient
- The generated ZIP file contains all the generated documents
- Image replacement requires the image paths in Excel to actually exist; supported formats are png / jpg / jpeg / gif / bmp
- Animated GIFs only show their first frame inside Word documents

## FAQ

**Q: What should I do if PDF export fails?**

A: Please check whether Microsoft Word or WPS Office is installed, and make sure Word documents can be opened normally.

**Q: The generated document format is incorrect?**

A: Please check whether the Word template file format is correct and whether the placeholders match the Excel headers.

**Q: How do I customize templates?**

A: In a Word document, use markers such as XXX01, XXX02 as placeholders. These markers will be replaced with the data from the corresponding Excel columns.

**Q: How do I use the multi-option feature?**

A:

1. Use field names starting with `check` in Excel (e.g., `check01`)
2. In the cell, separate multiple options with commas, Chinese commas, or the enumeration comma (e.g., `Swimming,Running,Reading`)
3. In the Word template, wrap the checkbox options with `{{check01}}` (e.g., `{{check01□ Swimming □ Running □ Reading}}`)
4. The system will automatically mark the selected options with `☑`

**Q: Which separators are supported for multi-options?**

A: The English comma `,`, the Chinese comma `，`, and the enumeration comma `、` are supported as option separators.

**Q: What should I do if checkboxes are not checked correctly?**

A: Please check:

1. Whether the option names in Excel match the option text in the Word template
2. Whether the correct separators are used
3. Whether the checkboxes in the Word template use the `□` symbol

**Q: How do I use the image replacement feature?**

A:

1. Use field names starting with `img` in Excel (e.g., `img01`)
2. Fill the cell with an image file path — absolute paths (e.g., `D:\pics\a.jpg`) and relative paths (e.g., `photos/a.jpg`, resolved against the Excel file's directory) are both supported
3. Write the same name (e.g., `img01`) in the Word template where the image should be inserted; the placeholder should preferably occupy its own paragraph
4. During generation the placeholder is replaced with the image; if the file is missing or the format is unsupported, the position is left empty and a warning is shown

**Q: What should I do if an image is not replaced (position left empty)?**

A: Please check:

1. Whether the path in the Excel cell actually exists (relative paths are resolved against the Excel file's directory)
2. Whether the image format is png / jpg / jpeg / gif / bmp
3. Whether the placeholder in the Word template matches the Excel header
