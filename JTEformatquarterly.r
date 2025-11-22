#JTE is close pending answers from email...

source("M:\\PAH\\Transplant Analyst Data\\duyXQ\\scripts\\outRealm\\ETL-R-main\\sourcery.R")

txepi <- importXLsheets(r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\TransplantEpisodeRecords.xlsx)') 
txepis <- txepi$Sheet1 %>%
  mutate(across(ends_with("ID"), as.character)) %>%
  mutate(MRN = as.character(`Medical Record NBR`))
str(txepis)

provid <- importXLsheets(r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\ProviderRecords.xlsx)') 
provids <- provid$Sheet1 %>%
  mutate(across(ends_with("ID"), as.character))
str(provids)

insuranceid <- importXLsheets(r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\InsuranceRecords.xlsx)') 
insuranceids <- insuranceid$Sheet1 %>%
  mutate(across(ends_with("ID"), as.character))
str(insuranceids)

dhx <- importXLsheets(r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\DialysisRecords.xlsx)') 
dhxs <- dhx$Sheet1 %>%
  mutate(across(ends_with("ID"), as.character)) %>%
  filter(!is.na(DialysisStartDTS))

str(dhxs)

rawout <- txepis %>%
  left_join(provids, by = c("Referring Provider ID" = "Provider ID")) %>%
  left_join(insuranceids, by = c("MRN" = "Patient MRN ID")) %>%
  left_join(dhxs, by = c("MRN" = "PatientMRNID"))

str(rawout)
rawout$`Transplant Evaluation End DTS` <- coalesce(rawout$`Transplant Evaluation End DTS` ,rawout$`Committee Review DTS`)


#clipr::write_clip(rawout)

df <- rawout %>%
  group_by(`Transplant Episode Record ID`, `Transplant Referral DS`) %>%
  arrange(desc(`Member Effective From DS`),desc(`DialysisStartDTS`), is.na(`DialysisEndDTS`), `DialysisEndDTS`) %>%
  slice(1) %>%
  ungroup()

str(df)
#clipr::write_clip(df)


df2 <- df %>%
	mutate(race = `Patient Race NM`,
	ethnicity = `Patient Ethinc Group NM`,
	sex = `Patient Sex NM`)

# make new col then case
#case the race
table(df2$race)
df2$race[grepl("American Indian or Alaska Native",df2$race)] <- "1"
df2$race[grepl("Asian",df2$race)] <- "2"
df2$race[grepl("Chinese",df2$race)] <- "2"
df2$race[grepl("Asian Indian",df2$race)] <- "2"

df2$race[grepl("Black or African American",df2$race)] <- "3"
df2$race[grepl("Pacific Islander",df2$race)] <- "4"
df2$race[grepl("Native Hawaiian",df2$race)] <- "4"
df2$race[grepl("White or Caucasian",df2$race)] <- "5"
# df2$race[grepl("Multi-Racial",df2$race)] <- "6"
df2$race[grepl("Other Race",df2$race)] <- "7"

df2$race[grepl("Patient Refused/Unknown",df2$race)] <- "7"
df2$race[grepl("Unknown",df2$race)] <- "7"
df2$race[is.na(df2$race) | df2$race == ""] <- "7"
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
df2$sex[is.na(df2$sex) | df2$sex == ""] <- "3"

table(df2$sex)

# case did_patient_dialyze handled before
# table(df2$did_patient_dialyze)
# df2$did_patient_dialyze[is.na(df2$did_patient_dialyze)] <- "0"
# df2$did_patient_dialyze[grepl("logi",df2$did_patient_dialyze)] <- "0"
# df2$did_patient_dialyze[grepl("Yes",df2$did_patient_dialyze)] <- "1"
# table(df2$did_patient_dialyze)

# check zip and fac name if dialyzed

#needs case to bring over current phase if past eval for eval close reason
df3 <- df2 %>%
  mutate(
    pre_dialysis = DialysisStartDTS - `Transplant Referral DS`,
    pre_dialysis_flag = case_when(
      pre_dialysis < 0 ~ "0",
      pre_dialysis >= 0 ~ "1"
    ),
    referral_closure_reason = case_when(
      `Current Phase Of Transplant NM` == "Referral" ~ `Current Reason For Transplant Phase And Status NM`
    ),
    eval_closure_reason = case_when(
      `Current Phase Of Transplant NM` == "Evaluation" ~ `Current Reason For Transplant Phase And Status NM`
    ),
	dialysis_facility_address = paste(DialysisCenterAddressLine1TXT, DialysisCenterAddressLine2TXT, sep = " "
	),
	referral_closure_date = case_when(
      `Current Phase Of Transplant NM` != "Referral" ~ `Translant Evaluation Start DS`,
	  `Current Phase Of Transplant NM` == "Referral" ~ `Last Transplant Phase Update DTS`
	),
  ) %>%
  filter(!is.na(MRN) & MRN != "") %>%
  mutate(DZipCode = substr(DialysisCenterZipCD, 1, 5))
  

df4 <- df3 %>%
	select(c(MRN,race,ethnicity,sex,"Transplant Referral DS","Translant Evaluation Start DS","Transplant Evaluation End DTS","Center Specific Waitlist DTS","pre_dialysis_flag","DialysisStartDTS","DialysisCenterNM","DialysisCenterMedicareIDNBR",dialysis_facility_address
	,DialysisCenterCityNM,"DialysisCenterStateNM","DZipCode",referral_closure_date, referral_closure_reason,eval_closure_reason,"Current Phase Of Transplant NM","Current Phase Of Transplant Status NM","Current Reason For Transplant Phase And Status NM","Last Transplant Phase Update DTS"))
write_xlsx(df4, r'(M:\PAH\Transplant Analyst Data\Transplant Coalition\jtemanualReasonsandsomedates2.xlsx)')


