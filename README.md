# ZIP TXT to Excel Converter

An R-based automation tool that converts multiple TXT files stored in ZIP
archives into Excel workbooks while preserving the source directory structure.

## Features

- Automatically discovers ZIP archives in source subdirectories.
- Extracts TXT files from each ZIP archive.
- Converts each ZIP archive into a separate Excel workbook.
- Creates one worksheet for each TXT file.
- Preserves the source folder structure in the output directory.
- Validates the number of worksheets in each workbook.
- Removes temporary extraction files after successful validation.

## Requirements

- R
- R package: `openxlsx`

Install the required package:

```r
install.packages("openxlsx")
```

## Directory Structure

```text
zip-txt-to-excel/
├── 01_r/
│   └── convert_txt_to_xlsx.R
├── 02_excel/
├── 03_data/
│   ├── Folder_A/
│   │   ├── sample1.zip
│   │   └── sample2.zip
│   └── Folder_B/
│       └── sample3.zip
├── .gitignore
└── README.md
```

## Usage

1. Open the project in RStudio.
2. Place ZIP archives inside subfolders of `03_data/`.
3. Run the script from the project root directory:

```r
source("01_r/convert_txt_to_xlsx.R")
```

4. Find the generated Excel workbooks in `02_excel/`.

## Output Example

```text
02_excel/
├── Folder_A/
│   ├── sample1.xlsx
│   └── sample2.xlsx
└── Folder_B/
    └── sample3.xlsx
```

Each TXT file becomes a separate worksheet within the corresponding Excel
workbook.

## Assumptions and Limitations

- TXT files are assumed to be tab-delimited.
- The first row of each TXT file is treated as column headers.
- Worksheet names are limited to 31 characters.
- Excel worksheet size limits apply.
- Validation checks file existence and worksheet counts, not complete data
integrity.
- Existing Excel output files with matching names are overwritten.

## Data Privacy

Original research data and generated Excel workbooks are excluded from version
control through `.gitignore`.