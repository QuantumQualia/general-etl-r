# radiant formatting and JTE has some extra steps
# do access steps first
source("M:\\PAH\\Transplant Analyst Data\\duyXQ\\scripts\\outRealm\\ETL-R-main\\sourcery.R")

df1 <- importXLsheets(r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\TxEpisodeData_Export2025.xlsx)')
str(df1)

# check col types nvm CSV is just strings

# perform logic on dialyze col using dates before we format them
df3 <- df1$TxEpisodeData_Export
oc <- df3$dialysis_start_date 
nc <- df3$referral_date
# make some true vals
dialyze <- oc < nc
table(!is.na(dialyze))
#select true vales from column and assign only those
df3$did_patient_dialyze[!is.na(dialyze)] <- "1"
df3$did_patient_dialyze[is.na(df3$did_patient_dialyze)] <- "0"


#replace new df this first time once, to prevent multiple
df2$did_patient_dialyze <- df3$did_patient_dialyze


#convert dates to MM/DD/YYYY formatting - if needed dates should be posix already
df2 <- df1$TxEpisodeData_Export
df2$date_of_birth <- strftime(df2$date_of_birth,"%m/%d/%Y",tz = "UTC")
df2$referral_date  <- strftime(df2$referral_date ,"%m/%d/%Y",tz = "UTC")
df2$evaluation_start_date <- strftime(df2$evaluation_start_date ,"%m/%d/%Y",tz = "UTC")
df2$evaluation_completion_date <- strftime(df2$evaluation_completion_date ,"%m/%d/%Y",tz = "UTC")
df2$waitlisting_date <- strftime(df2$waitlisting_date ,"%m/%d/%Y",tz = "UTC")
df2$dialysis_start_date <- strftime(df2$dialysis_start_date ,"%m/%d/%Y",tz = "UTC")

# fix SSN to 000-00-0000 only

df2$social_security_number[grepl("999-99-9999",df2$social_security_number)] <- "000-00-0000"
df2$social_security_number[grepl("777-77-7777",df2$social_security_number)] <- "000-00-0000"
df2$social_security_number[is.na(df2$social_security_number)] <- "000-00-0000"
# checks
df2$social_security_number[grepl("999-99-9999",df2$social_security_number)] 
df2$social_security_number[grepl("777-77-7777",df2$social_security_number)]
df2$social_security_number[is.na(df2$social_security_number)]

#case the race
table(df2$race)
df2$race[grepl("American Indian or Alaska Native",df2$race)] <- "1"
df2$race[grepl("Asian",df2$race)] <- "2"
df2$race[grepl("Black or African American",df2$race)] <- "3"
df2$race[grepl("Pacific Islander",df2$race)] <- "4"
df2$race[grepl("Native Hawaiian",df2$race)] <- "4"
df2$race[grepl("White or Caucasian",df2$race)] <- "5"
# df2$race[grepl("Multi-Racial",df2$race)] <- "6"
df2$race[grepl("Other Race",df2$race)] <- "7"
# we dont classify Asian Indians in radiant so i made them other
df2$race[grepl("Asian Indian",df2$race)] <- "7"

df2$race[grepl("Patient Refused/Unknown",df2$race)] <- "8"
df2$race[grepl("Unknown",df2$race)] <- "8"

table(df2$race)

# case ethnicity
table(df2$ethnicity)
df2$ethnicity[grepl("Hispanic or Latino",df2$ethnicity)] <- "1"
df2$ethnicity[grepl("Not Hispanic",df2$ethnicity)] <- "2"
df2$ethnicity[grepl("Black",df2$ethnicity)] <- "2"
df2$ethnicity[grepl("White",df2$ethnicity)] <- "2"
df2$ethnicity[grepl("Patient Refused/Unknown",df2$ethnicity)] <- "3"
df2$ethnicity[is.na(df2$ethnicity)] <- "3"
table(df2$ethnicity)

# case sex
table(df2$sex)
df2$sex[grepl("Female",df2$sex)] <- "1"
df2$sex[grepl("Male",df2$sex)] <- "2"
df2$sex[grepl("Unknown",df2$sex)] <- "3"
table(df2$sex)

# case did_patient_dialyze handled before
# table(df2$did_patient_dialyze)
# df2$did_patient_dialyze[is.na(df2$did_patient_dialyze)] <- "0"
# df2$did_patient_dialyze[grepl("logi",df2$did_patient_dialyze)] <- "0"
# df2$did_patient_dialyze[grepl("Yes",df2$did_patient_dialyze)] <- "1"
# table(df2$did_patient_dialyze)

# check zip and fac name if dialyzed
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_name) )
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_ccn) )
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_address) )
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_city) )
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_state) )
subset(df2, did_patient_dialyze == "1" & is.na(dialysis_facility_zip_code) )
 
write_xlsx(df2, r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\PED_ReferralData_01.2024 Thru_12.2024 Datawithmissingzips2.xlsx)')


