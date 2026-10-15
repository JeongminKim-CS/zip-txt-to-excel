# ============================================================
# Script: convert_txt_to_xlsx.R
# ============================================================
#
# Description:
#   This script converts TXT files stored inside ZIP archives
#   into Excel workbooks.
#
#   Each ZIP archive is converted into one Excel workbook,
#   with each TXT file stored in a separate worksheet.
#
# Input:
#   - ZIP files located in:
#     03_data/<folder_name>/
#
# Output:
#   - Excel workbooks saved in:
#     02_excel/<folder_name>/
#
# Processing Steps:
#   1. Identify subfolders in 03_data.
#   2. Locate ZIP archives in each source subfolder.
#   3. Extract TXT files from each ZIP archive.
#   4. Read TXT files as tabular data.
#   5. Create one Excel workbook per ZIP archive.
#   6. Store each TXT file in a separate worksheet.
#   7. Save the workbook in the corresponding output folder.
#
# Naming Convention:
#   - Excel filename: Same as the source ZIP filename.
#   - Worksheet name: Based on the source TXT filename.
#   - Output folder: Same as the source folder name.
#
# Dependencies:
#   - openxlsx
#
# Notes:
#   - TXT files are assumed to be tab-delimited.
#   - The first row is assumed to contain column headers.
#   - Excel worksheet names are limited to 31 characters.
#   - Temporary extracted files are removed after processing.
#
# ============================================================

# ============================================================
# 0. Load Required Packages
# ============================================================

library(openxlsx)

# ============================================================
# 1. Configuration
# ============================================================

# Define input and output directories
input_dir <- "03_data"
output_dir <- "02_excel"

# Check whether the input directory exists
if ( !dir.exists(input_dir)){
  stop("Input directory does not exist: ",
       input_dir)
}

# Create output directory if it does not exist
if (!dir.exists(output_dir)){
  dir.create(
    output_dir,
    recursive = TRUE
  )
}

# Display configured directories
cat("Input directory: ", normalizePath(input_dir), "\n")
cat("Output directory:", normalizePath(output_dir), "\n")


# ============================================================
# 2. Folder Discovery
# ============================================================

# Identify immediate subdirectories in the input directory
source_folders <- list.dirs(
  path = input_dir,
  full.names = TRUE,
  recursive = FALSE
)

# Check whether source folders were found
if (length(source_folders) == 0) {
  stop("No subdirectories found in: ",
       input_dir)
}

# Display discovered folders
cat("Number of source folders:", length(source_folders), "\n")

print(source_folders)

# ============================================================
# 3. ZIP File Discovery
# ============================================================

# Initialize a list to store ZIP file paths by source folder
zip_file_list <- list()

# Search ZIP files in each source folder
for ( source_folder in source_folders) {
  
  # get source folder name
  folder_name <- basename(source_folder)
  
  # Define ZIP directory
  zip_dir <- source_folder
  
  # Check whether ZIP directory exists
  if (!dir.exists(zip_dir)) {
    warning("ZIP directory not found: ", zip_dir)
    next
  }
  
  # Find ZIP files
  zip_files <- list.files(
    path = zip_dir,
    pattern = "\\.zip$",
    full.names = TRUE,
    ignore.case = TRUE
  )
  
  # Store ZIP file paths
  zip_file_list[[folder_name]] <- zip_files
  
  # Display results
  cat("\nSource folder:", folder_name, "\n")
  cat("ZIP files found:", length(zip_files), "\n")
  
  print(basename(zip_files))
}


# ============================================================
# 4. TXT File Discovery & Extraction
# ============================================================

# Initialize a list to store TXT file information
txt_file_list <- list()

# Initialize a list to track temporary directories
temp_dir_list <- list()

# Process each source folder
for (folder_name in names(zip_file_list)) {
  
  txt_file_list[[folder_name]] <- list()
  
  # Process each ZIP archive
  for (zip_file in zip_file_list[[folder_name]]) {
    
    # Get ZIP filename without extension
    zip_name <- tools::file_path_sans_ext(
      basename(zip_file)
    )
    
    # List files inside ZIP archive
    zip_contents <- unzip(
      zip_file,
      list = TRUE
    )
    
    # Identify valid TXT files
    txt_files <- zip_contents$Name[
      grepl("\\.txt$", zip_contents$Name, ignore.case = TRUE) &
        !grepl("(^|/)__MACOSX/", zip_contents$Name) &
        !grepl("(^|/)\\._", zip_contents$Name)
    ]
    
    # Skip ZIP archives without TXT files
    if (length(txt_files) == 0) {
      warning("No TXT files found in: ", zip_file)
      next
    }
    
    # Create a temporary extraction directory
    temp_dir <- tempfile()
    dir.create(temp_dir)
    
    # Track temporary directory for cleanup
    temp_dir_list[[temp_dir]] <- temp_dir
    
    # Extract TXT files
    unzip(
      zipfile = zip_file,
      files = txt_files,
      exdir = temp_dir
    )
    
    # Store extracted TXT file paths
    extracted_files <- file.path(temp_dir, txt_files)
    
    txt_file_list[[folder_name]][[zip_name]] <- extracted_files
    
    # Display results
    cat("\nSource folder:", folder_name, "\n")
    cat("ZIP archive:", basename(zip_file), "\n")
    cat("TXT files found:", length(txt_files), "\n")
    
    print(basename(txt_files))
  }
}




# ============================================================
# 5. TXT Import & Excel Export
# ============================================================

# Process each source folder
for (folder_name in names(txt_file_list)) {
  
  # Define output directory
  excel_dir <- file.path(output_dir, folder_name)
  
  # Create output directory if it does not exist
  dir.create(
    excel_dir,
    recursive = TRUE,
    showWarnings = FALSE
  )
  
  # Process each ZIP archive
  for (zip_name in names(txt_file_list[[folder_name]])) {
    
    # Create a new Excel workbook
    wb <- createWorkbook()
    
    # Get extracted TXT files
    txt_files <- txt_file_list[[folder_name]][[zip_name]]
    
    # Process each TXT file
    for (txt_file in txt_files) {
      
      # Read tab-delimited TXT file
      data <- read.delim(
        file = txt_file,
        header = TRUE,
        check.names = FALSE
      )
      
      # Generate worksheet name
      sheet_name <- tools::file_path_sans_ext(
        basename(txt_file)
      )
      
      # Remove invalid characters
      sheet_name <- gsub(
        "[\\[\\]\\*\\?/\\\\:]",
        "_",
        sheet_name
      )
      
      # Limit worksheet name to 31 characters
      sheet_name <- substr(sheet_name, 1, 31)
      
      # Ensure unique worksheet names
      base_name <- sheet_name
      i <- 1
      
      while (tolower(sheet_name) %in% tolower(names(wb))) {
        suffix <- paste0("_", i)
        
        sheet_name <- paste0(
          substr(base_name, 1, 31 - nchar(suffix)),
          suffix
        )
        
        i <- i + 1
      }
      
      # Add worksheet
      addWorksheet(
        wb,
        sheetName = sheet_name
      )
      
      # Write TXT data to worksheet
      writeData(
        wb,
        sheet = sheet_name,
        x = data
      )
    }
    
    # Define Excel output file
    output_file <- file.path(
      excel_dir,
      paste0(zip_name, ".xlsx")
    )
    
    # Save Excel workbook
    saveWorkbook(
      wb,
      file = output_file,
      overwrite = TRUE
    )
    
    # Display progress
    cat(
      "Created:", output_file,
      "| Sheets:", length(names(wb)),
      "\n"
    )
  }
}


# ============================================================
# 6. Cleanup & Validation
# ============================================================

# Initialize validation results
validation_results <- data.frame(
  Folder = character(),
  ZIP = character(),
  Expected_Sheets = integer(),
  Actual_Sheets = integer(),
  Status = character()
)

# Validate each converted ZIP archive
for (folder_name in names(txt_file_list)) {
  
  for (zip_name in names(txt_file_list[[folder_name]])) {
    
    # Define Excel file path
    excel_file <- file.path(
      output_dir,
      folder_name,
      paste0(zip_name, ".xlsx")
    )
    
    # Expected number of worksheets
    expected <- length(
      txt_file_list[[folder_name]][[zip_name]]
    )
    
    # Check Excel file and worksheets
    if (file.exists(excel_file)) {
      
      actual <- length(getSheetNames(excel_file))
      
      status <- if (expected == actual) {
        "PASS"
      } else {
        "FAIL"
      }
      
    } else {
      actual <- 0L
      status <- "MISSING"
    }
    
    # Record validation result
    validation_results <- rbind(
      validation_results,
      data.frame(
        Folder = folder_name,
        ZIP = zip_name,
        Expected_Sheets = expected,
        Actual_Sheets = actual,
        Status = status
      )
    )
  }
}

# Display validation summary
print(validation_results)

cat(
  "\nPassed:",
  sum(validation_results$Status == "PASS"),
  "\nFailed:",
  sum(validation_results$Status != "PASS"),
  "\n"
)


# ============================================================
# 6.1. Temporary File Cleanup
# ============================================================

# Check whether all conversions passed validation
all_passed <- nrow(validation_results) > 0 &&
  all(validation_results$Status == "PASS")

if (all_passed) {
  
  # Get temporary directory paths
  temp_dirs <- unlist(
    temp_dir_list,
    use.names = FALSE
  )
  
  # Remove temporary directories
  for (temp_dir in temp_dirs) {
    
    if (dir.exists(temp_dir)) {
      unlink(
        temp_dir,
        recursive = TRUE
      )
    }
  }
  
  # Verify cleanup
  remaining_dirs <- temp_dirs[
    dir.exists(temp_dirs)
  ]
  
  cleaned_count <- length(temp_dirs) - length(remaining_dirs)
  
  if (length(remaining_dirs) == 0) {
    message("All temporary directories removed.")
  } else {
    warning(
      "Some temporary directories could not be removed."
    )
  }
  
  cat(
    "\nTemporary directories cleaned:",
    cleaned_count, "/", length(temp_dirs), "\n"
  )
  
} else {
  
  warning(
    "Validation failed. Temporary files were preserved."
  )
}

message("Validation and cleanup completed.")


