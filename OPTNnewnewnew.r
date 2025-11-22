#OPTN this is the real one
#ALUA was missing because it was too high up on the data, childrens is now getting cuttoff and thats ok


# OPTN Vols
source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\sourcery.R)')
# source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\OPTNnewnewnew.R)')


#TXP 
library(readxl)
library(dplyr)
options(dplyr.summarise.inform = FALSE)
library(tidyr)
library(purrr)

shell.exec('https://optn.transplant.hrsa.gov/data/view-data-reports/build-advanced/')
readline('(txp settings\ncategory- Transplant\nreport cols- organ, donor type\nrows- transplant center\ndisplay- counts, portrait\nthen choose current year\nexport to folder but data just needs to be copied into the newest excel sheet for the report\n)')

readline('(WL settings\ncategory- Waiting list additions\nreport cols- organ\nrows- transplant center\ndisplay- counts, portrait\nthen choose current year)')
readline('(copy previous report into new file, rn appropriately.\n Add new data(to premade Tx and WL sheets) to the correct months)')
readline('(If you are doing this for a new year, update all the years in all the sheets\n replace all current data with "0" for months that havent happend yet before starting)')
readline('(Close the sheet, the script cant read it if its already open.)')
readline('(Note the sheet will look for the most recent excel file in the dir and process it)')
readline('(The last steps will be to paste into the Tx and WL sheets respectively witht he updated monthly data)')
readline('(Enter to Start, all you need to do is make sure the input file is correct)')

# this does a lot....
process_sheet <- function(file_path, sheet_name) {
  # column headers from row 2 and centers from row 3
  col_headers <- read_excel(file_path, sheet = sheet_name, range = "A2:P2", col_names = FALSE)
  center_headers <- read_excel(file_path, sheet = sheet_name, range = "A3:P3", col_names = FALSE)
  
  data <- read_excel(file_path, sheet = sheet_name, skip = 4)
  
  # rn columns dynamically
  names(data) <- as.character(col_headers[1,])
  
  # Find column indices for specific data
  kidney_all_idx <- which(col_headers[1,] == "Kidney" & center_headers[1,] == "All Donor Types")
  liver_all_idx <- which(col_headers[1,] == "Liver" & center_headers[1,] == "All Donor Types")
  pancreas_all_idx <- which(col_headers[1,] == "Pancreas" & center_headers[1,] == "All Donor Types")
  kidney_pancreas_all_idx <- which(col_headers[1,] == "Kidney / Pancreas" & center_headers[1,] == "All Donor Types")
  heart_all_idx <- which(col_headers[1,] == "Heart" & center_headers[1,] == "All Donor Types")
  
  # rename columns
result <- data %>%
  select(
    Center = 1,  # First column is always Centers
    Kidney = all_of(kidney_all_idx),
    Liver = all_of(liver_all_idx),
    Pancreas = all_of(pancreas_all_idx),
    `Kidney / Pancreas` = all_of(kidney_pancreas_all_idx),
    Heart = all_of(heart_all_idx)
  ) %>%
  filter(Center != "All Centers")
  
  return(result)
}

# Read all sheets
#file_path <- r'(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\OPTNvolumes\CY2024WL update_only_VolumesJan.xlsx)'
QRattendfiles<- list.files(path = r"(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\OPTNvolumes)" ,pattern = "TxpWL", full.names = TRUE, ignore.case = TRUE)
file_info <- file.info(QRattendfiles)
file_info$file_name <- rownames(file_info)
sorted_files <- QRattendfiles[order(desc(file_info$ctime))]
file_path<-sorted_files[1]

sheet_names <- excel_sheets(file_path)
sheet_names <- sheet_names[grepl("\\d{4} Tx$", sheet_names)]  # looks for "#### Tx"


all_data <- lapply(sheet_names, function(sheet) {
  data <- process_sheet(file_path, sheet)
  data$Month <- sub(" Tx$", "", sheet)
  data
})

combined_data2 <- bind_rows(all_data)


combined_data3 <- combined_data2 %>%
  pivot_wider(names_from = Month, values_from = c(Kidney, Liver, Pancreas, `Kidney / Pancreas`, Heart)) %>%
  # Create temporary columns with original cumulative values
  mutate(
    across(ends_with("Jan") | ends_with("Feb") | ends_with("March") | ends_with("Apr") | 
           ends_with("May") | ends_with("June") | ends_with("July") | 
           ends_with("Aug") | ends_with("Sept") | ends_with("Oct") | 
           ends_with("Nov") | ends_with("Dec"), 
           ~., .names = "{.col}_temp")
  ) %>%
  # Calculate differences using temporary columns
  mutate(
    across(ends_with("Feb"), ~ . - get(paste0(sub("Feb$", "Jan", cur_column()), "_temp"))),
    across(ends_with("March"), ~ . - get(paste0(sub("March$", "Feb", cur_column()), "_temp"))),
    across(ends_with("Apr"), ~ . - get(paste0(sub("Apr$", "March", cur_column()), "_temp"))),
    across(ends_with("May"), ~ . - get(paste0(sub("May$", "Apr", cur_column()), "_temp"))),
    across(ends_with("June"), ~ . - get(paste0(sub("June$", "May", cur_column()), "_temp"))),
    across(ends_with("July"), ~ . - get(paste0(sub("July$", "June", cur_column()), "_temp"))),
    across(ends_with("Aug"), ~ . - get(paste0(sub("Aug$", "July", cur_column()), "_temp"))),
    across(ends_with("Sept"), ~ . - get(paste0(sub("Sept$", "Aug", cur_column()), "_temp"))),
    across(ends_with("Oct"), ~ . - get(paste0(sub("Oct$", "Sept", cur_column()), "_temp"))),
    across(ends_with("Nov"), ~ . - get(paste0(sub("Nov$", "Oct", cur_column()), "_temp"))),
	across(ends_with("Dec"), ~ . - get(paste0(sub("Dec$", "Nov", cur_column()), "_temp")))
  ) %>%
  
  select(-ends_with("_temp"))
# If you want to pivot back to the original format
combined_data <- combined_data3 %>%
  pivot_longer(cols = -Center, 
               names_to = c(".value", "Month"), 
               names_pattern = "(.+)_(.+)") %>%
  mutate(Month = word(Month, 1))

# totals for each center and organ
organ_totals <- combined_data %>%
  group_by(Center) %>%
  summarise(across(Kidney:`Kidney / Pancreas`:Heart, sum, na.rm = TRUE))

# Unified function for both TX and WL data
get_top_centers <- function(data, organ, n = 10) {
  required_centers <- c("GAPH-TX1 Piedmont Hospital", "ALUA-TX1 University of Alabama Hospital", "GAEM-TX1 Emory University Hospital", 
                       "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
                       "FLSL-TX1 Mayo Clinic Hospital Florida")
# =OR($B4="ALUA-TX1 University of Alabama Hospital",$B4="GAEM-TX1 Emory University Hospital",$B4="TNVU-TX1 Vanderbilt University Medical Center",$B4="SCLA-TX1 MUSC Lancaster",$B4="FLSL-TX1 Mayo Clinic Hospital Florida")
  # Get top centers excluding required ones
  top_non_required <- data %>%
    filter(!Center %in% required_centers) %>%
    arrange(desc(!!sym(organ))) %>%
    slice_head(n = n) %>%
    pull(Center)
  
  # Get required centers' actual values
  required_in_data <- data %>%
    filter(Center %in% required_centers) %>%
    pull(Center)
  
  unique(c(required_in_data, top_non_required))[1:(n + length(required_centers))]
}

get_top_centers <- function(data, organ, n = 10) {
  # List of centers to always include
  required_centers <- c("GAPH-TX1 Piedmont Hospital", "UAB", "GAEM-TX1 Emory University Hospital", 
                        "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
                        "FLSL-TX1 Mayo Clinic Hospital Florida", "ALUA-TX1 University of Alabama Hospital")
  
  # Get required centers that exist in the data
  existing_required <- intersect(required_centers, data$Center)
  
  # Get top n centers excluding required ones
  top_centers <- data %>%
    filter(!Center %in% existing_required) %>%
    arrange(desc(!!sym(organ))) %>%
    slice_head(n = n) %>%
    pull(Center)
  
  # Combine and ensure required centers come first
  all_centers <- unique(c(existing_required, top_centers))
  
  # Limit to total desired length (n + number of required centers)
  all_centers[1:min(length(all_centers), n + length(existing_required))]
}


# Get top centers for each organ
organs <- c("Kidney", "Liver", "Pancreas", "Kidney / Pancreas", "Heart")
top_centers_by_organ <- lapply(organs, function(organ) get_top_centers(organ_totals, organ))
names(top_centers_by_organ) <- organs


generate_organ_output <- function(data, organ_name) {
  # Define required centers
  required_centers <- c(
    "GAPH-TX1 Piedmont Hospital", "UAB", "GAEM-TX1 Emory University Hospital", 
    "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
    "FLSL-TX1 Mayo Clinic Hospital Florida", "ALUA-TX1 University of Alabama Hospital"
  )
  
  # Get top centers (guaranteed to include all required)
  top_centers <- get_top_centers(organ_totals, organ_name)
  
  # Get previous month and year
  report_date <- Sys.Date() %m-% months(1)
  previous_month <- month(report_date)
  previous_year <- year(report_date)
  
  # Month column names (adjust if your data uses different names)
  month_columns <- c("Jan", "Feb", "March", "Apr", "May", "June",
                     "July", "Aug", "Sept", "Oct", "Nov", "Dec")
  previous_month_name <- month_columns[previous_month]
  
  # Pivot data to wide format
  wide_data <- data %>%
    filter(Center %in% top_centers) %>%
    select(Center, !!sym(organ_name), Month) %>%
    pivot_wider(names_from = Month, values_from = !!sym(organ_name), values_fill = 0)
  
  # Ensure all required centers are present, add rows of zeros if missing
  missing_centers <- setdiff(required_centers, wide_data$Center)
  if (length(missing_centers) > 0) {
    # Create a tibble with zeros for missing centers
    zero_row <- as.list(rep(0, length(month_columns)))
    names(zero_row) <- month_columns
    missing_df <- tibble(
      Center = missing_centers
    ) %>%
      bind_cols(as_tibble(zero_row))
    wide_data <- bind_rows(wide_data, missing_df)
  }
  
  # Join with totals
  wide_data <- wide_data %>%
    left_join(organ_totals %>% select(Center, Total = !!sym(organ_name)), by = "Center")
  
  # Arrange by Total
  wide_data <- wide_data %>%
    arrange(desc(Total)) %>%
    mutate(
      Ranking = row_number(),
      CYTD = if (previous_month_name %in% names(.)) !!sym(previous_month_name) else NA_real_,
      .after = Center
    )
  
  # Select columns in desired order
  wide_data <- wide_data %>%
    select(Ranking, Center, all_of(month_columns), CYTD, Total)
  
  # Print missing required centers by month
  for (center in required_centers) {
    if (center %in% wide_data$Center) {
      center_row <- wide_data %>% filter(Center == center)
      missing_months <- month_columns[center_row[1, month_columns] == 0]
      if (length(missing_months) > 0) {
        cat(sprintf("Center '%s' has no data for months: %s\n", center, paste(missing_months, collapse = ", ")))
      }
    } else {
      cat(sprintf("Center '%s' is completely missing from the data.\n", center))
    }
  }
  
  return(wide_data)
}
# Generate output for each organ
organ_outputs <- lapply(organs, function(organ) generate_organ_output(combined_data, organ))
names(organ_outputs) <- organs


organ_outputs <- map(names(organ_outputs), function(df_name) {
  organ_outputs[[df_name]] %>%
    rename(!!df_name := Center)
}) %>%
  set_names(names(organ_outputs))
  
  

###
###recalc months because its complicated
###
# Function 1: Calculate month-specific values from cumulative data
calculate_monthly_values <- function(df) {
  # Make a copy of the dataframe
  result_df <- df
  
  # Identify all month columns (only those that are actual calendar months)
  month_names <- c("Jan", "Feb", "March", "Apr", "May", "June", 
                   "July", "Aug", "Sept", "Oct", "Nov", "Dec")
  
  # Find which month columns actually exist in the dataframe
  month_cols <- intersect(names(df), month_names)
  
  # Convert month columns to numeric
  for (col in month_cols) {
    result_df[[col]] <- as.numeric(as.character(result_df[[col]]))
  }
  
  # Find which months have data (non-zero values in any row)
  months_with_data <- c()
  for (col in month_cols) {
    if (any(result_df[[col]] > 0, na.rm = TRUE)) {
      months_with_data <- c(months_with_data, col)
    }
  }
  
  # Order months chronologically
  months_with_data <- month_names[month_names %in% months_with_data]
  
  # Process each month except the first one
  if(length(months_with_data) > 1) {
    for(i in 2:length(months_with_data)) {
      current_month <- months_with_data[i]
      prev_months <- months_with_data[1:(i-1)]
      
      # Subtract the sum of all previous months from the current month
      result_df[[current_month]] <- result_df[[current_month]] - rowSums(result_df[prev_months], na.rm = TRUE)
    }
  }
  
  return(result_df)
}


recalculate_cytd <- function(df) {
  # Create modified copy without altering original
  result_df <- df
  
  # Define valid month columns based on your actual data
  month_names <- c("Jan", "Feb", "March", "Apr", "May", "June", 
                   "July", "Aug", "Sept", "Oct", "Nov", "Dec")
  
  # Find existing month columns
  month_cols <- intersect(names(result_df), month_names)
  
  # Identify months with actual data (non-zero values)
  months_with_data <- month_cols[sapply(result_df[month_cols], 
                                      function(x) any(x > 0, na.rm = TRUE))]
  
  # Order months chronologically
  months_with_data <- month_names[month_names %in% months_with_data]
  
if("CYTD" %in% names(result_df)) {
  # Ensure CYTD is numeric and handle NA values
  result_df$CYTD <- as.numeric(result_df$CYTD)
  
  # Sort by CYTD descending, then by Center name ascending
  result_df <- result_df %>%
    arrange(desc(CYTD))
  
  # Reset row names for cleanliness
  rownames(result_df) <- NULL
}
  
  return(result_df)
}


# Function to process all organs
process_all_organs <- function(organ_list) {
  # First calculate monthly values for all organs
  monthly_values <- lapply(organ_list, calculate_monthly_values)
  
  # Then recalculate CYTD for all organs
  processed_organs <- lapply(monthly_values, recalculate_cytd)
  
  return(processed_organs)
}

# Example usage with a single dataframe:
# processed_data <- calculate_monthly_values(df)
# processed_data <- recalculate_cytd(processed_data)

# For the full list:
organ_outputs2 <- process_all_organs(organ_outputs)
# monthly_values <- lapply(organ_outputs, calculate_monthly_values)

# Paste individual organ outputs, highlight GAPH in sheet.
# clipr::write_clip(organ_outputs$Kidney)
# clipr::write_clip(organ_outputs$`Kidney / Pancreas`)
# clipr::write_clip(organ_outputs$Liver)
# clipr::write_clip(organ_outputs$Pancreas)
# clipr::write_clip(organ_outputs$Heart)

combined_dftxp <- map_df(names(organ_outputs2), function(df_name) {
  df <- organ_outputs2[[df_name]]
  col_names <- names(df)
  new_df <- rbind(col_names, df)
  new_df <- mutate(new_df, Organ = df_name)
  return(new_df)
}) %>%
  select(Organ, everything())

#paste into sheet and shift organ/center names col to left and re-sort the organs
#shell.exec(file_path)
clipr::write_clip(combined_dftxp)


##################
#################
## WL part borrows combined data, organ total


library(readxl)
library(dplyr)
library(tidyr)
library(purrr)

QRattendfiles<- list.files(path = r"(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\OPTNvolumes)" ,pattern = "TxpWL", full.names = TRUE, ignore.case = TRUE)
file_info <- file.info(QRattendfiles)
file_info$file_name <- rownames(file_info)
sorted_files <- QRattendfiles[order(desc(file_info$ctime))]
file_path<-sorted_files[1]

# Updated process_sheet function for new structure
process_sheetWL <- function(file_path, sheet_name) {
  # Read data starting from row 1 (headers are different now)
  data <- read_excel(file_path, sheet = sheet_name, col_names = FALSE)
  
  # Extract column names from the third row (where organ names are)
  col_names <- data[2, ]
  
  # Read actual data starting from row 3
  data <- data[3:nrow(data), ]
  
  # Assign column names
  names(data) <- as.character(unlist(col_names))
  # Replace NA or blank name for the first column with "Center"
  names(data)[1] <- "Center"
  
  # Select and rename columns
  result <- data %>%
    select(
      Center,  # First column is always Centers
      Kidney,
      Liver,
      Pancreas,
      `Kidney / Pancreas`,
      Heart
    ) %>%
    filter(Center != "All Centers")
  col_names
  return(result)
}
# Read all sheets
# file_path <- r'(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\OPTNvolumes\CY2024WL update_only_VolumesJan.xlsx)'  # Update with your file path
sheet_names <- excel_sheets(file_path)
sheet_names <- sheet_names[grepl("\\d{4} WL$", sheet_names)]

all_data <- lapply(sheet_names, function(sheet) {
  data <- process_sheetWL(file_path, sheet)
  data$Month <- sub(" WL$", "", sheet)
  data
})

# Combine all sheets
combined_data2 <- bind_rows(all_data)

# Ensure numeric values before pivoting
combined_data2 <- combined_data2 %>%
  mutate(across(-c(Center, Month), ~as.numeric(as.character(.))))

combined_data3 <- combined_data2 %>%
  # First, pivot the data wider
  pivot_wider(names_from = Month, values_from = c(Kidney, Liver, Pancreas, `Kidney / Pancreas`, Heart)) %>%
  # Ensure all value columns are numeric
  mutate(across(-Center, ~as.numeric(as.character(.)))) %>%
  # Create temporary columns
  mutate(
    across(ends_with("Jan") | ends_with("Feb") | ends_with("March") | ends_with("Apr") | 
           ends_with("May") | ends_with("June") | ends_with("July") | 
           ends_with("Aug") | ends_with("Sept") | ends_with("Oct") | 
           ends_with("Nov") | ends_with("Dec"), 
           ~as.numeric(.), .names = "{.col}_temp")
  ) %>%
  # Calculate differences
  mutate(
    across(ends_with("Feb"), ~as.numeric(.) - as.numeric(get(paste0(sub("Feb$", "Jan", cur_column()), "_temp")))),
    across(ends_with("March"), ~as.numeric(.) - as.numeric(get(paste0(sub("March$", "Feb", cur_column()), "_temp")))),
    across(ends_with("Apr"), ~as.numeric(.) - as.numeric(get(paste0(sub("Apr$", "March", cur_column()), "_temp")))),
    across(ends_with("May"), ~as.numeric(.) - as.numeric(get(paste0(sub("May$", "Apr", cur_column()), "_temp")))),
    across(ends_with("June"), ~as.numeric(.) - as.numeric(get(paste0(sub("June$", "May", cur_column()), "_temp")))),
    across(ends_with("July"), ~as.numeric(.) - as.numeric(get(paste0(sub("July$", "June", cur_column()), "_temp")))),
    across(ends_with("Aug"), ~as.numeric(.) - as.numeric(get(paste0(sub("Aug$", "July", cur_column()), "_temp")))),
    across(ends_with("Sept"), ~as.numeric(.) - as.numeric(get(paste0(sub("Sept$", "Aug", cur_column()), "_temp")))),
    across(ends_with("Oct"), ~as.numeric(.) - as.numeric(get(paste0(sub("Oct$", "Sept", cur_column()), "_temp")))),
    across(ends_with("Nov"), ~as.numeric(.) - as.numeric(get(paste0(sub("Nov$", "Oct", cur_column()), "_temp")))),
    across(ends_with("Dec"), ~as.numeric(.) - as.numeric(get(paste0(sub("Dec$", "Nov", cur_column()), "_temp"))))
  ) %>%
  # Remove temporary columns
  select(-ends_with("_temp"))
# If you want to pivot back to the original format
combined_data <- combined_data3 %>%
  pivot_longer(cols = -Center, 
               names_to = c(".value", "Month"), 
               names_pattern = "(.+)_(.+)") %>%
  mutate(Month = word(Month, 1))
# Calculate totals for each center and organ
organ_totals <- combined_data %>%
  group_by(Center) %>%
  summarise(across(Kidney:`Kidney / Pancreas`:Heart, sum, na.rm = TRUE))

# Unified function for both TX and WL data
get_top_centers <- function(data, organ, n = 10) {
  required_centers <- c("GAPH-TX1 Piedmont Hospital", "ALUA-TX1 University of Alabama Hospital", "GAEM-TX1 Emory University Hospital", 
                       "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
                       "FLSL-TX1 Mayo Clinic Hospital Florida")
  
  # Get top centers excluding required ones
  top_non_required <- data %>%
    filter(!Center %in% required_centers) %>%
    arrange(desc(!!sym(organ))) %>%
    slice_head(n = n) %>%
    pull(Center)
  
  # Get required centers' actual values
  required_in_data <- data %>%
    filter(Center %in% required_centers) %>%
    pull(Center)
  
  unique(c(required_in_data, top_non_required))[1:(n + length(required_centers))]
}

get_top_centers <- function(data, organ, n = 10) {
  required_centers <- c("GAPH-TX1 Piedmont Hospital", "ALUA-TX1 University of Alabama Hospital", "GAEM-TX1 Emory University Hospital", 
                       "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
                       "FLSL-TX1 Mayo Clinic Hospital Florida")
  
  # Get top centers excluding required ones
  top_non_required <- data %>%
    filter(!Center %in% required_centers) %>%
    arrange(desc(!!sym(organ))) %>%
    slice_head(n = n) %>%
    pull(Center)
  
  # Get required centers' actual values
  required_in_data <- data %>%
    filter(Center %in% required_centers) %>%
    pull(Center)
  
  # Combine required centers and top non-required centers
  top_centers <- unique(c(required_in_data, top_non_required))[1:(n + length(required_centers))]
  
  # Check for missing required centers
  missing_centers <- setdiff(required_centers, required_in_data)
  if (length(missing_centers) > 0) {
    warning("The following required centers are missing from the data: ", paste(missing_centers, collapse = ", "))
  }
  
  return(top_centers)
}


# Get top centers for each organ
organs <- c("Kidney", "Liver", "Pancreas", "Kidney / Pancreas", "Heart")
top_centers_by_organ <- lapply(organs, function(organ) get_top_centers(organ_totals, organ))
names(top_centers_by_organ) <- organs

# Generate organ output with guaranteed required centers
generate_organ_output <- function(data, organ_name) {
  top_centers <- get_top_centers(data, organ_name)  # Use modified function
  
  data %>%
    filter(Center %in% top_centers) %>%
    select(Center, !!sym(organ_name), Month) %>%
    pivot_wider(names_from = Month, values_from = !!sym(organ_name), values_fill = 0) %>%
    left_join(data %>% select(Center, Total = !!sym(organ_name)), by = "Center") %>%
    arrange(desc(Total)) %>%
    mutate(
      Ranking = row_number(),
      CYTD = rowSums(select(., Jan:Dec), na.rm = TRUE)
    ) %>%
    select(Ranking, Center, Jan:Dec, CYTD, Total)
}




generate_organ_output <- function(data, organ_name) {
  # Define required centers
  required_centers <- c(
    "GAPH-TX1 Piedmont Hospital", "UAB", "GAEM-TX1 Emory University Hospital", 
    "TNVU-TX1 Vanderbilt University Medical Center", "SCLA-TX1 MUSC Lancaster", 
    "FLSL-TX1 Mayo Clinic Hospital Florida", "ALUA-TX1 University of Alabama Hospital"
  )
  
  # Get top centers (guaranteed to include all required)
  top_centers <- get_top_centers(organ_totals, organ_name)
  
  # Get previous month and year
  report_date <- Sys.Date() %m-% months(1)
  previous_month <- month(report_date)
  previous_year <- year(report_date)
  
  # Month column names (adjust if your data uses different names)
  month_columns <- c("Jan", "Feb", "March", "Apr", "May", "June",
                     "July", "Aug", "Sept", "Oct", "Nov", "Dec")
  previous_month_name <- month_columns[previous_month]
  
  # Pivot data to wide format
  wide_data <- data %>%
    filter(Center %in% top_centers) %>%
    select(Center, !!sym(organ_name), Month) %>%
    pivot_wider(names_from = Month, values_from = !!sym(organ_name), values_fill = 0)
  
  # Ensure all required centers are present, add rows of zeros if missing
  missing_centers <- setdiff(required_centers, wide_data$Center)
  if (length(missing_centers) > 0) {
    # Create a tibble with zeros for missing centers
    zero_row <- as.list(rep(0, length(month_columns)))
    names(zero_row) <- month_columns
    missing_df <- tibble(
      Center = missing_centers
    ) %>%
      bind_cols(as_tibble(zero_row))
    wide_data <- bind_rows(wide_data, missing_df)
  }
  
  # Join with totals
  wide_data <- wide_data %>%
    left_join(organ_totals %>% select(Center, Total = !!sym(organ_name)), by = "Center")
  
  # Arrange by Total
  wide_data <- wide_data %>%
    arrange(desc(Total)) %>%
    mutate(
      Ranking = row_number(),
      CYTD = if (previous_month_name %in% names(.)) !!sym(previous_month_name) else NA_real_,
      .after = Center
    )
  
  # Select columns in desired order
  wide_data <- wide_data %>%
    select(Ranking, Center, all_of(month_columns), CYTD, Total)
  
  # Print missing required centers by month
  for (center in required_centers) {
    if (center %in% wide_data$Center) {
      center_row <- wide_data %>% filter(Center == center)
      missing_months <- month_columns[center_row[1, month_columns] == 0]
      if (length(missing_months) > 0) {
        cat(sprintf("Center '%s' has no data for months: %s\n", center, paste(missing_months, collapse = ", ")))
      }
    } else {
      cat(sprintf("Center '%s' is completely missing from the data.\n", center))
    }
  }
  
  return(wide_data)
}





# Generate output for each organ
organ_outputs <- lapply(organs, function(organ) generate_organ_output(combined_data, organ))
names(organ_outputs) <- organs

organ_outputs <- map(names(organ_outputs), function(df_name) {
  organ_outputs[[df_name]] %>%
    rename(!!df_name := Center)
}) %>%
  set_names(names(organ_outputs))

names(organ_outputs$Kidney[2])[names(organ_outputs$Kidney[2]) == "Center"] <- "Kidney"
names(organ_outputs$Kidney[2]) <- "Kidney"
names(organ_outputs$`Kidney / Pancreas`[2]) <- "Kidney / Pancreas"
names(organ_outputs$Liver[2]) <-  "Liver"
names(organ_outputs$Pancreas[2]) <- "Pancreas"
names(organ_outputs$Heart[2]) <- "Heart"

organ_outputs2 <- process_all_organs(organ_outputs)

# Paste individual organ outputs, highlight in sheet, 
# clipr::write_clip(organ_outputs$Kidney)
# clipr::write_clip(organ_outputs$`Kidney / Pancreas`)
# clipr::write_clip(organ_outputs$Liver)
# clipr::write_clip(organ_outputs$Pancreas)
# clipr::write_clip(organ_outputs$Heart)

combined_dfWL <- map_df(names(organ_outputs2), function(df_name) {
  df <- organ_outputs2[[df_name]]
  col_names <- names(df)
  new_df <- rbind(col_names, df)
  new_df <- mutate(new_df, Organ = df_name)
  return(new_df)
}) %>%
  select(Organ, everything())

organ_order <- c("Kidney", "Kidney / Pancreas", "Liver", "Pancreas","Heart")
#sort txp and WL
combined_dftxp$Organ <- factor(combined_dftxp$Organ, levels = organ_order)
combined_dftxp <- combined_dftxp[order(combined_dftxp$Organ), ]

combined_dfWL$Organ <- factor(combined_dfWL$Organ, levels = organ_order)
combined_dfWL <- combined_dfWL[order(combined_dfWL$Organ), ]

## final formatting, NA for UAB for easy selecting
combined_dftxp2 <- combined_dftxp %>% mutate(Kidney = coalesce(!!!lapply(select(., (which(names(.) == "Total") + 1):ncol(.)), function(x) ifelse(x == "", NA, x)), Kidney)) %>%
  mutate(Kidney = if_else(Kidney == "UAB", "", as.character(Kidney)))

combined_dfWL2 <- combined_dfWL %>% mutate(Kidney = coalesce(!!!lapply(select(., (which(names(.) == "Total") + 1):ncol(.)), function(x) ifelse(x == "", NA, x)), Kidney))  %>%
  mutate(Kidney = if_else(Kidney == "UAB", "", as.character(Kidney)))
	
#paste into sheet and shift organ/center names col to left and resort the organs
shell.exec(file_path)
clipr::write_clip(combined_dftxp2)
readline(r'(Tx copied to system clipboard, excel sheet should be opened.)')
readline(r'(paste into main Tx sheet (on the right side of the report)\n -- and shift organ/center names col to left\n -- re-sort the organs to match current report order)')
readline(r'(once done with Tx, press Enter and WL will be in clipboard)')
clipr::write_clip(combined_dfWL2)
readline(r'( WL in clipboard paste into main WL sheet and resort organs like Tx)')
readline(r'( Enter to open email for sending)')
toddy <- Sys.Date()
shell.exec(str_glue(r'(mailto:Leah.Baker@piedmont.org;Clark.Kensinger@piedmont.org;Eric.Gibney@piedmont.org;Sundus.Lodhi@piedmont.org;Ezequiel.Molina@piedmont.org;Jonathan.Hundley@piedmont.org;Elizabeth.Miller1@piedmont.org?subject=OPTN Volume {toddy} &cc=Teandra.Lassiter@piedmont.org;Desriee.Plummer@piedmont.org&body=Attached.)'))


source_dir <- r'(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\OPTNvolumes)'
dest_dir <- r'(M:\PAH\Transplant Quality\25. OPTNvolumes)'

xlsx_files <- dir_info(source_dir, glob = "*.xlsx") %>%
  arrange(desc(modification_time))

if (nrow(xlsx_files) > 0) {
  file_copy(xlsx_files$path[1], dest_dir)
  cat("Copied", basename(xlsx_files$path[1]), "to", dest_dir)
} else {
  cat("No xlsx files found in", source_dir)
}

readline(r'(copied to M:\PAH\Transplant Quality\25. OPTNvolumes\n)')
