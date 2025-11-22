# this doesnt automate yet, emails are in antibodybatch.r - home file might need MRN col reverted
# Get old dialysis data, put current dialysis batch orders into new batch order folder first
# get data out of tableau for the sheets, episode, history, 
# overview, paste in home, paste in dialysis - they are created
## check for dupes, add new centers, make sure the hotfix addresses are updated
# update home kit number
# fix the addresses with PID
# update missing PID in Dialysis, searching on biotouch is preferred
source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\sourcery.R)')

todaydir<- Sys.Date()
old_datad <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\New Batch Orders)", pattern = "Dialysis", full.names = TRUE, ignore.case = TRUE)

old_dataf <- file.info(old_datad)
old_dataf$file_name <- rownames(old_dataf)

sorted_filesod <- old_datad[order(desc(old_dataf$ctime))]

oldd1 <- sorted_filesod[1]
oldd1

oldd12 <- importXLsheets(oldd1)
str(oldd12)

oldd13 <- oldd12$Sheet1
oldd13$`Patient ID (MRN/EMPI)` <- as.character(oldd13$`Patient ID (MRN/EMPI)`)
str(oldd13)

excel_files <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\New Batch Orders)", 
                         pattern = "\\.xlsx$", 
                         full.names = TRUE)

combined_data <- lapply(excel_files, function(file) {
  read_excel(file, sheet = 1) %>% 
    mutate(Quantity = as.character(Quantity),
	Zipcode = as.character(Zipcode),
			`Patient ID (MRN/EMPI)` = as.character(`Patient ID (MRN/EMPI)`),
           SourceFile = basename(file))
}) %>% 
  bind_rows()

distinct_odata <- combined_data %>% 
  distinct(`Patient ID (MRN/EMPI)`, `Client ID`, .keep_all = TRUE)

# make dialysis center list key with pids from previous report
ngdata <- unique(oldd13[c('Client ID', 'Ship To Name', 'Provider/Lab.', 'Contact', 'Address 1', 'Address 2', 'City', 'ST', 'Zipcode', 'Country', 'Phone .')])
str(ngdata)

#make current home mrns to not resend To
hexcel_files <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch Orders)", 
                         pattern = ".*Home.*\\.xlsx$", 
                         full.names = TRUE)

hexcel_files

#home output needs col renamed
hcombined_data <- lapply(hexcel_files, function(file) {
  read_excel(file, sheet = 1) %>% 
    mutate(
      Quantity = as.character(Quantity),
      `Patient ID (MRN/EMPI)` = as.character(`Patient ID (MRN/EMPI)`),
      `Client ID` = as.character(`Client ID`),  
      SourceFile = basename(file)
    )
}) %>% 
  bind_rows()


hdistinct_odata2 <- hcombined_data %>% 
  distinct(`Client ID`) %>%
  as.data.frame()
hdistinct_odata <- hdistinct_odata2$`Client ID`
str(hdistinct_odata)




#get waitlist active
txpepisodesraw <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders)", pattern = "TransplantEpisode", full.names = TRUE, ignore.case = TRUE)

file_infot <- file.info(txpepisodesraw)
file_infot$file_name <- rownames(file_infot)

sorted_filest <- txpepisodesraw[order(desc(file_infot$ctime))]

qrAttenddt <- sorted_filest[1]
qrAttenddt

qrAttenddt2 <- importXLsheets(qrAttenddt)
str(qrAttenddt2)

qrAttenddt3 <- qrAttenddt2$Sheet1

qrAttenddt3$PatientMRNID <- as.character(qrAttenddt3$`Medical Record NBR`)
str(qrAttenddt3)

allactandinactWL <- qrAttenddt3 %>% 
  filter(str_detect(tolower(`Current Phase Of Transplant Status NM`), "active")) %>%
    filter(!(
    `Current Phase Of Transplant Status NM` == "Inactive" &
    as.Date(Sys.Date()) - as.Date(`Last Transplant Phase Update DTS`) > 365
  ))
# import txp history for names
txpepisodepth <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders)", pattern = "TransplantEpisodes", full.names = TRUE, ignore.case = TRUE)

file_infote <- file.info(txpepisodepth)
file_infote$file_name <- rownames(file_infote)

sorted_fileste <- txpepisodepth[order(desc(file_infote$ctime))]

qrAttendte <- sorted_fileste[1]
qrAttendte

txpepi <- importXLsheets(qrAttendte)
str(txpepi)

reltxpepi <- txpepi$Sheet1 %>%
  distinct(`Medical Record NBR`, .keep_all = TRUE) %>% 
  select(-c("Transplant Episode Record ID","Patient NM","Tx Episode Organ UNOS Definition","Transplant Episode Organ Name","Current Phase Of Transplant NM","Current Phase Of Transplant Status NM","Current Phase Of Transplant Status ID","Patient Home Phone NBR","Last Transplant Phase Update DTS"))
 
 #current patient info
 str(reltxpepi)

# get epic dialysis history file
dhist <- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders)", pattern = "DialysisHistory", full.names = TRUE, ignore.case = TRUE)

file_infod <- file.info(dhist)
file_infod$file_name <- rownames(file_infod)

sorted_filesd <- dhist[order(desc(file_infod$ctime))]

qrAttendd <- sorted_filesd[1]
qrAttendd

qrAttendd2 <- importXLsheets(qrAttendd)
str(qrAttendd2)

qrAttendd33 <- qrAttendd2$Sheet1 

qrAttendd33$PatientMRNID <- as.character(qrAttendd33$PatientMRNID)

qrAttendd3 <- qrAttendd33 %>%
	left_join(reltxpepi, by =c("PatientMRNID" = "Medical Record NBR"))

str(qrAttendd3)

#dialysis patients with center data only
final_data2 <- qrAttendd3 %>% 
  filter(!str_detect(tolower(DialysisTypeNM), "home"))%>% 
  filter(!DialysisCenterID == 10)


final_data1 <- subset(final_data2,final_data2$PatientMRNID %in% allactandinactWL$PatientMRNID)
length(unique(final_data1$PatientMRNID))
nrow(final_data1)


# patients with home, dont need center data for these
homedialy <- qrAttendd3 %>% 
  filter(str_detect(tolower(DialysisTypeNM), "home"))

#rel home patients without dialysis
activinactivehomenoHD <- subset(allactandinactWL,!(allactandinactWL$`Medical Record NBR` %in% qrAttendd3$PatientMRNID))
str(activinactivehomenoHD)
# get home hemo from episodes
homedialyaddr <- subset(allactandinactWL,allactandinactWL$`Medical Record NBR` %in% homedialy$PatientMRNID)
str(homedialyaddr)

allhomedialy2 <- rbind(activinactivehomenoHD,homedialyaddr)
length(unique(allhomedialy2$PatientMRNID)) == nrow(allhomedialy2)

#get dupes for joyce
homedialyaddr <- subset(allactandinactWL,allactandinactWL$PatientMRNID%in% homedialy$PatientMRNID) %>%
  group_by(`Medical Record NBR`) %>%
  filter(n() > 1) %>%
  ungroup()
  
nrow(homedialyaddr)


#active folks - periotineal = dialysis + home - alllll the hemo folks which are a lot and might have been in the old report...
nrow(allhomedialy2)+ length(unique(final_data1$PatientMRNID))
nrow(allactandinactWL) 

allhomedialy <- subset(allhomedialy2,!(allhomedialy2$`Medical Record NBR` %in% hdistinct_odata))
nrow(allhomedialy)
# full kit change to dialy2 below
cleanDhome <- allhomedialy %>% select(c("Medical Record NBR", "Patient NM","Patient Address Line 1 TXT","Patient Address Line 2 TXT","Patient City NM","Patient State NM","Patient Zip CD","Patient Home Phone NBR","PatientMRNID","Patient Social Security NBR","Patient Last NM", "Patient First NM","Patient Birth DTS", )) %>%
  mutate(`Provider/Lab` = NA, .after = `Patient NM`)%>%
  mutate(`Contact` = NA, .after = `Patient NM`)%>%
  mutate(`Country` = "US", .after = `Patient Zip CD`)%>%
  mutate(`Sex` = NA, .after = `Patient Birth DTS`)%>%
  mutate(`Race` = NA, .after = `Sex`)%>%
  mutate(`Phys` = NA, .after = `Race`)%>%
  mutate(`Diagnosis` = NA, .after = `Phys`)%>%
  mutate(`WL` = NA, .after = `Diagnosis`)%>%
  mutate(`ICD-9` = NA, .after = `WL`)%>%
  mutate(`Item .` = "1891P", .after = `ICD-9`)%>%
  mutate(`Quantity` = NA, .after = `Item .`)%>%
  mutate(`Shipping Service Level` = "Ground", .after = `Quantity`)
  
rownames(cleanDhome) <- NULL
# Sample DataFrame
df <- cleanDhome

# List of IDs to remove
#mrn_to_remove <- c("911506566","4")
mrn_to_remove <- c("4")

removed_rows <- df %>% filter(PatientMRNID %in% mrn_to_remove) %>% select(c("PatientMRNID","Patient NM"))
print("Removed rows:")
print(removed_rows)

df_cleaned <- df %>% filter(!PatientMRNID %in% mrn_to_remove)

print("Cleaned DataFrame:")
print(df_cleaned)
clipr::write_clip(cleanDhome)
cat('cleaned home in clipboard check')
str(cleanDhome)
write_xlsx(cleanDhome, str_glue(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch orders\BatchOrderHome{todaydir}.xlsx)'))
##################################
###PASTE #####################################
###################################

# give pid to dialysis folks
final_dataD <- final_data1 %>%
  left_join(distinct_odata, by = c("DialysisCenterAddressLine1TXT" = "Address 1"), relationship = "many-to-many") %>%
  left_join(distinct_odata, by = c("PatientMRNID" = "Patient ID (MRN/EMPI)" )) %>%
  mutate(across(
    ends_with(".x"), 
    ~ coalesce(., get(sub(".x$", ".y", cur_column()))),
    .names = "{.col}"
  )) %>%
  group_by(PatientMRNID) %>%
  distinct(PatientMRNID, .keep_all = TRUE) %>% 
  ungroup() 
  
clipr::write_clip(final_dataD)

cleanDial <- final_dataD %>% select(c("Client ID.x", "DialysisCenterNM","Provider/Lab..x","Contact.x","DialysisCenterAddressLine1TXT","DialysisCenterAddressLine2TXT","DialysisCenterCityNM","DialysisCenterStateNM","DialysisCenterZipCD","Country.x","DialysisCenterPhoneNBR","PatientMRNID","Patient Social Security NBR","Patient Last NM", "Patient First NM","Patient Birth DTS", "Sex.x", "Race.x", "Physician.x", "Diagnosis.x", "Wait List.x", "ICD-9.x", "Item ..x", "Quantity.x", "Shipping Service Level.x", "Comments.x"))
clipr::write_clip(cleanDial %>% select(1:26))
str(cleanDial)

cat('cleaned dialysis1 in clipboard before old addresses are edited')
write_xlsx(cleanDial, str_glue(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch orders\BatchOrderDialysispre{todaydir}.xlsx)'))
#format colums

shell.exec('mailto:Evan.Matthews@piedmont.org;Joyce.Anderson@piedmont.org;Kristi.Goldberg@piedmont.org?subject=Antibody Batch Kits Uploaded&body=Uploaded and Attached Antibody Kit Batch Upload files.')
#check for special cases below

#start new orders for newly listed since last batch?
#key MRN, status, last update and use that to filter from episodes list
## also add to adhoc file



# Select specific columns for the final dataset
#finald <- final_dataD %>% select(c("Client ID.x", "Client ID.y", "Ship To Name.x", "Ship To Name.x", "DialysisCenterNM", "Ship To Name.y", "Provider/Lab..x", "Contact.x", "Address 1", "DialysisCenterAddressLine1TXT", "Address 2.x", "Address 2.y", "DialysisCenterAddressLine2TXT", "City.x", "City.y", "ST.x", "ST.y", "Zipcode.x", "Zipcode.y", "Country.x", "Country.y", "Phone ..x", "Phone ..y", "Patient ID (MRN/EMPI)", "Social Security.", "Patient Last Name", "Patient First Name", "DOB", "Sex", "Race", "Physician", "Diagnosis", "Wait List", "ICD-9", "Item .", "Quantity", "Shipping Service Level", "Comments"))

#finald <- final_dataD %>% select(c("Client ID", "Ship To Name", "DialysisCenterNM", "Provider/Lab.", "Contact", "DialysisCenterAddressLine1TXT", "DialysisCenterAddressLine2TXT","DialysisCenterCityNM", "DialysisCenterStateNM", "Address 2", "City.x", "City.y", "ST.x", "ST.y", "Zipcode.x", "Zipcode.y", "Country.x", "Country.y", "Phone ..x", "Phone ..y", "Patient ID (MRN/EMPI)", "Social Security.", "Patient Last Name", "Patient First Name", "DOB", "Sex", "Race", "Physician", "Diagnosis", "Wait List", "ICD-9", "Item .", "Quantity", "Shipping Service Level", "Comments"))

# clipr::write_clip(finald)


#remove home orders in progress since 6



#get completed center sheet
readline("Import newly created dialysis history and add formatting")
QRattendfiles<- list.files(path = r"(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch orders)" ,pattern = "Dialysispre", full.names = TRUE, ignore.case = TRUE)
file_info <- file.info(QRattendfiles)
file_info$file_name <- rownames(file_info)
sorted_files <- QRattendfiles[order(desc(file_info$ctime))]
qrAttend<-sorted_files[1]
qrAttend
# print(sorted_files)
#joining attendance sheets
qrAttend2 <- importXLsheets(qrAttend)
qrAttend3 <- qrAttend2$Sheet1
# qrAttend3 <- qrAttend2$BatchOrderTemplate1 %>% subset(., `Client ID`)
str(qrAttend3)

davitajonesid <- "PD1027"
crimsonid <- "PD1678"
dsikenn <- "PD1333"
#maybe update phone check against spectra
spaldingcountyid <- "PD1157"



qrAttend3jofix <- qrAttend3 %>%
  rename(`Client ID` = `Client ID.x`) %>%
  mutate(
    `DialysisCenterAddressLine1TXT` = case_when(
      !is.na(`Client ID`) & `Client ID` == davitajonesid ~ "1595 Stockbridge Road",
      !is.na(`Client ID`) & `Client ID` == spaldingcountyid ~ "744 S. 8th Street",
      TRUE ~ `DialysisCenterAddressLine1TXT`
    ),
    `DialysisCenterPhoneNBR` = case_when(
      !is.na(`Client ID`) & `Client ID` == davitajonesid ~ "678-833-1921",
      !is.na(`Client ID`) & `Client ID` == crimsonid ~ "205-752-3267",
      !is.na(`Client ID`) & `Client ID` == dsikenn ~ "678-354-7124",
      TRUE ~ `DialysisCenterPhoneNBR`
    )
  )


clipr::write_clip(qrAttend3jofix)
readline('cleaned dialysis2final in clipboard')
write_xlsx(qrAttend3jofix, str_glue(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch orders\BatchOrderDialysis{todaydir}.xlsx)'))
readline('fill in blank PIDs and add new centers if needed')
#need to filter out adhoc ignores from XL file