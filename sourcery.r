library(dplyr)
library(readr)
library(readxl)
library(writexl)
library(explore)
library(dplyr)
library(openxlsx)
library(clipr)
library(tidyr)
library(stringr)
library(lubridate)
# library(arrow)
# library(tidyverse)
# library(tictoc)
# library(RODBC) #tried... not simple but tableaau -> access -> excel seems fast also would help if you needed a midway backup
# library(odbc)
# library(DBI)
# con2 <- dbConnect(odbc(), .connection_string = paste0("Driver={SQL Server};Dbq=",'M:\\PAH\\Transplant Objective Measures\\FY2024 Objective Measures\\Kidney\\txpexport1.mdb'), ";")
# srcpath = r('M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\WLeligiblelodhi.r')
# gsub("\\", "/", srcpath, fixed = TRUE)
# # source(srcpath)
# srcpath

# normalizePath('M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\WLeligiblelodhi.r')
cat("useful funs

rmnonprintall - for rm non printing chars in anything
next fun - descrip
next fun - descrip
")


# def gsubs and def controller func to rm all currently known non printing chars
rmnonprintgsub1 <- function(x) {
  gsub("\u00A0", "", x)
}

rmnonprintgsub2 <- function(x) {
  gsub("\u00C2", "", x)
}

rmnonprintall <- function(your_df) {
your_df <- your_df %>%
mutate_all(rmnonprintgsub1) %>%
mutate_all(rmnonprintgsub2)
return(your_df)
}

convertXLnumbdate <- function(your_vec) {
XLcdate<- convertToDate(your_vec, origin = "1900-01-01")
return(XLcdate)
}

#usage
#df415Rjoin$Transplant.Surgery.DTS <- convertextdatemmddyyyy(df415Rjoin$Transplant.Surgery.DTS)
convertextdatemmddyyyy <- function(your_vec) {
textdate22 <- as.Date(your_vec,format="%m/%d/%Y",tz=Sys.timezone())
return(textdate22)
}

is.POSIXct <- function(x) inherits(x, "POSIXct")

convertextdatemmddyyyytime <- function(your_vec) {
textdate23 <- as.Date(your_vec,format="%m/%d/%Y %H:%M",tz=Sys.timezone())
return(textdate23)
}

#usage
#colvec <- c("X155.66.3732")
#getduperows(df415,colvec)
getduperows <- function(datafrm,colvec){
dupfromcolvec <- sort(datafrm[[colvec]][duplicated(datafrm[[colvec]]) | duplicated(datafrm[[colvec]], fromLast=TRUE)])
duperows <- subset(datafrm,datafrm[[colvec]] %in% dupfromcolvec) %>% arrange(colvec)
# fix?
# dupfromcolvec <- sort(unetASKclean$Center.Patient.Id[duplicated(unetASKclean$Center.Patient.Id)])
# duperows <- subset(unetASKclean,unetASKclean$Center.Patient.Id %in% dupfromcolvec) %>% arrange(Center.Patient.Id)


return(duperows)
}

# importing multiple sheets
# usage
# sheetpath <- "M:\\PAH\\Transplant Analyst Data\\duyXQ\\LiverQAPIALLMaster.xlsx"
# all_sheetsM <- importXLsheets(sheetpath)
importXLsheets <- function(sheetpath0){
sheet_names11 <- excel_sheets(sheetpath0)           # Get sheet names
# sheet_namesU
all_sheets11 <- lapply(sheet_names11, function(x) {          # Read all sheets to list
  as.data.frame(read_excel(sheetpath0, sheet = x)) } )
names(all_sheets11) <- sheet_names11
names(all_sheets11)
return(all_sheets11)
}

# usaage
# tockWLrem <- read_dataframe_from_clipboard()
read_dataframe_from_clipboard <- function() {
  # Read data from the clipboard
  df <- read.table("clipboard", header = TRUE, sep = "\t", stringsAsFactors = FALSE, quote="",fill = TRUE)
  return(df)
}

pickataskk <- function() {
cat('What do we want to do? \n 1. "Envelope and A2B2 Data export (simple export from Tableau)" \n 2. "Kidney Evals Due" \n 3. Thursday Report for Kristi \n 4. Daily CMS Reports \n 5. Weekly rmvals not ready \n 6. Adverse Events weekly validate XL file import due to extra cells being made active during creation \n 7. Heart VAD ECMO Vols \n 8. Heart Committee monthly [ASO] \n 9. Antibody batch kits\n 11. Referral monthly
\n 12. Meeting Attendance\n 13. Remaining at risk ALL Organs\n 14. RLs ACCESS -> Tableau steps\n 15. Status reports DialandNeph\n 16. Shirley Future ABO report - email\n\n')
your_vecx <- as.numeric(readline('Please choose an option by providing me with a number from the list above \n '))
if ( is.na(your_vecx)) {
print("not a choice bro")
} else if ( your_vecx == 1) {
	print("starting Envelope Export Task standby for Tableau load")
	envlopeanda2b2instructions()
} else if ( your_vecx == 2) {
	print("starting Kidney Evals Due")
	kidneyEvalsDue()
} else if ( your_vecx == 3) {
	print("starting Thursday Report")
	ThursdayReportInstructions()
} else if ( your_vecx == 4) {
	print("CMS Daily Mon+Tues")
	CMSdailyinstructions()
# } else if ( your_vecx == 5) {
	# print("Transplant rm Weekly audit")
	# TransplantRemovalinstructions()
} else if ( your_vecx == 6) {
	print("Adverse Events Every Friday")
	AdverseEvents()
} else if ( your_vecx == 7) {
	print("VAD ECMO Vols, check and reconnect data sources so the numbers update.")
	VadEcmoVolinstructions()
} else if ( your_vecx == 8) {
	print("Heart Committee Teandra [ASO]")
	MonthlyHeartCommittee()
} else if ( your_vecx == 9) {
	print("AntiBodyBatchKits")
	AntiBodyBatchKits()
} else if ( your_vecx == 11) {
	print("Referrals monthly, prob neeed to be on 35w")
	Referralsmonthly()
} else if ( your_vecx == 12) {
	print("meeting attendance")
	source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\meetingAttend.R)')
	getattendees()
} else if ( your_vecx == 13) {
	print("remaining At risk all organs")
	source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\remainingatriskdirectfromepic.R)')
	print("verify the 1YR results (SRTR Metrics Tab) with XynQAPI\n complete if you want to do some hunting not required\n")
	shell.exec(r'(https://xyn.my.site.com/XynQAPI/apex/QAHomePage)')
	print("all done after you send those emails")
} else if ( your_vecx == 14) {
	print("RLs Access -> Tableau")
	AdverseEventsAccesstoTableau()
} else if ( your_vecx == 15) {
	print("Status Dialysis/Neph Reports")
	source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\DialysisCenterReportmonthly.r)')
} else if ( your_vecx == 16) {
	print("Status Dialysis/Neph Reports")
	source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\futureABOapptshirley.r)')

} else {
print("final catch, this won't print")
}

return(your_vecx)
}

scrubHTMLcodefromtext <- function(htmlString) {
 htmlString <- gsub("<br>", " ",htmlString)
 htmlString <- gsub("\\n", "   ",htmlString)
 htmlString <- gsub("&nbsp;", " ",htmlString)
 htmlString <- gsub("<//div>", " ",htmlString)
 htmlString <- gsub("<div>", " ",htmlString)
 htmlString <- gsub("</?[^>]+>", " ",htmlString)
}

envlopeanda2b2instructions <- function() {
AccessDBDataforUploadTAB <- r'(M:\PAH\Transplant Analyst Data\AccessDBDataForUpload.twb)'
shell.exec(AccessDBDataforUploadTAB)
readline('Tableau should be up \n 1. "enter to proceed"\n')
readline('We are going to copy paste each sheet into a new file that will open \n "enter to proceed"\n')

readline('Copy to Crosstab "PatientsForEnvelope" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\PatientsEPICForEnvelopes.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Copy to Crosstab "ReferringDrsForEnvelopes" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\ProvidersEPICForEnvelopes.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Copy to Crosstab "DialysisCtrForEnvelopes" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\DialysisCtrEPICForEnvelopes.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Copy to Crosstab "A2_A2B_Data" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\KidneyTxEpisodesForA1A2bDB.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')
readline('Task Complete remove from calendar for tracking  \n "enter to proceed"\n')
# return(a2b2env)
}


kidneyEvalsDue <- function() {
KidneyEvalsdueTAB <- r'(M:\PAH\Transplant Analyst Data\KidneyEvaluationsDueForReEval.twb)'
shell.exec(KidneyEvalsdueTAB)
readline('Tableau should be up \n 1. "enter to proceed"\n')
readline('We are going to copy paste each sheet into a new file that will open \n "enter to proceed"\n')

readline('Copy to Crosstab "Waitlist" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\Evals_Joyce\EPICWaitlist.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Copy to Crosstab "PastEvalAppts" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\Evals_Joyce\EPIC_kidneyEvals.xlsx)')
readline('replace all data  \n "enter to proceed"\n')

readline('Copy to Crosstab "FutureEvalAppts_" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\Evals_Joyce\EPIC_FutureKidneyEvalAppts.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Copy to Crosstab "All Kid Tx Episodes" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\Evals_Joyce\All Kidney Episodes.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n "enter to proceed"\n')

readline('Open Access file to run macros \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Reports Misc\Kidney\TDrive2024\Evals_Joyce\KidneyReEvalsDue.accdb)')
readline('After macro runs (sometimes slow) nothing to do this is complete  \n "enter to proceed"\n')

readline('Task Complete remove from calendar for tracking  \n "enter to proceed"\n')
# return(a2b2env)
}

ThursdayReportInstructions  <- function() {
readline("please use 35w machine\n")
shell.exec(r'(C:\Users\246397\Desktop\Epic.lnk)')
readline('Login to Epic and run pb_kidney waitlistedpts v3\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report)')
readline('Put in folder that is opened - filename pb####\n "enter to proceed"\n')
readline('open it and paste into this CSV - press enter\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList_Weights.csv)')
readline('save that csv then copy and paste into this next EXCEL - fixes some formatting will fix later "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList_Weights_ForAccess.xlsx)')
readline('Reclass MRN/SSN to text fields if they arent already "enter to proceed"\n')
shell.exec(r'(https://auth.unos.org/Login?redirect_uri=https%3A%2F%2Fportal.unos.org)')
readline('Login -> WL icon -> Reports \n export both KidneyListedKDPI \n rename to include date \n AND Points Report \n Points params: allocation points 35-85% select all Blood, candidate status is both "enter to proceed"\n')

readline('Open new KidneyListedKDPI (Text File) report and paste (checking headers) into following EXCEL \n ORGAN, MRN, SSN need to be text "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyListedKDPI.xlsx)')
readline('open new points report\n reclass: ORGAN and SSN as text\n save as KidneyCandidatePointsReport (XL), over write the old points \n dont forget to rename the sheet to KidneyCandidatePontsReport "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\DataSourceForCPRAMELDScores.twb)')
cat('starting tableau for copy crosstab procedures\n')
readline('Tableau should be up \n 1. "enter to proceed"\n')
readline('We are going to copy paste each sheet into a new file that will open \n "enter to proceed"\n')

readline('Copy to Crosstab "Dialysis Type" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList_DialyisType.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n Reclass: as text MRN, SSN, ORGAN\n')

readline('Copy to Crosstab "FullEvalProtocols" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyFullEval_Protocol.xlsx)')
readline('Replace data  \n Reclass: as text MRN, ORGAN\n')

readline('Copy to Crosstab "KidneyWaitlist" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyTransplantEpisodeData.xlsx)')
readline('Paste under first row and check headers then delete new headers row  \n Reclass: as text MRN, SSN, ORGAN, UNOS ID prob\n')

cat('moving old files \n KidneywaitlistDataFullEvals_All, \n QData_ForTop50_ThursdayRept, \n QData_KDPI_ThurdaysRept')
shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneywaitlistDataFullEvals_All.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\QData_ForTop50_ThursdayRept.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\QData_KDPI_ThurdaysRept.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList.accdb)')
cat('starting access\n\n')
clipr::write_clip(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneywaitlistDataFullEvals_All.xlsx)')
readline('export 2 QDATA tables into the Kidney WL Report dir \n\n external QDATA = KidneywaitlistDataFullEvals_All for file name\nfile path in clipboard\n')
readline('then run 2 Qdata queries that will make 2 more Qdata tables to export')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\QData_KDPI_ThurdaysRept.xlsx)')
readline('Paste the opened QDATA KDPI and paste into first sheet Thursday Report to open')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Thursday Report.xlsx)')
readline('ithink you can do this without changing machines...\n')
readline('goback to 73w to do final steps\n run this over there \n ThursdayReport73w() in clip\n')
clipr::write_clip("ThursdayReport73w()")
}
ThursdayReport73w <- function() {

shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\CPRAs_MELDScores.twb)')
readline('Thursday report will open, paste pink sheet into second Excel sheet \n should come from Weekly full eval active tab in Tableau')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Thursday Report.xlsx)')
readline('Save and email to Kristi :] you did it yay!')
toddy <- Sys.Date()
shell.exec(str_glue(r'(mailto:Kristi.Goldberg@piedmont.org?subject=Thursday Report {toddy} &body=Thursday Report Attached)'))
readline('Task Complete remove from calendar for tracking  \n "enter to proceed"\n')
# return(a2b2env)
}

copysubsetscms <- function(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2) {
rawepichr3 <- head(rawepichr2)
rawepickk3 <- head(rawepickk2)
rawepicli3 <- head(rawepicli2)
rawepicpanc3 <- head(rawepicpanc2)

# list of vars
targetorganscms <-(list(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2))
templateorganscms <-(list(rawepichr3,rawepickk3,rawepicli3,rawepicpanc3))

# list of inputs
mappednames <- map2(targetorganscms,templateorganscms, ~names(.y))
mapped1strow <- map2(targetorganscms,templateorganscms, ~.y[1,])

# switch title/headers

for (i in 1:length(targetorganscms)){
	names(targetorganscms[[i]]) <- mapped1strow[[i]]
	# print(i)
}
for (i in 1:length(targetorganscms)){
	targetorganscms[[i]][1,] <- mappednames[[i]] #char to date error here for some reason
	# print (i)
}


# rawepichr2[1,] <- names(rawepichr3)
# names(rawepichr2) <- rawepichr3[1,]

return(targetorganscms)
}

CMSdailyinstructions  <- function() {

todayfoldernm <- readline('enter dir for reports eg 03.20.2024: \n')
todaydir<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm})')
# check if sub directory exists 
if (dir.exists(todaydir)){
	# specifying the working directory
	cat('todaydir exists')
	shell.exec(todaydir)
} else {
	# create a new sub directory inside
	dir.create(file.path(todaydir))
	shell.exec(todaydir)
}
shell.exec(r'(C:\Users\246397\Desktop\Epic.lnk)')
clipr::write_clip(todaydir)
setwd(todaydir)
print(todaydir)
todaydirraw<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm}\raw)')
todaydirprint<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm}\printme)')
readline('Login to Epic and run \n 1. CMS Currently Admitted \n 2. Transplanted 18 months \n (path already in clipboard if needed)\n')
readline('3. DAR - one for MTP and one for heart \n report made but if needed to build...\n change visit type to list post Tx follow, living donor follow\n')
readline('Patient lists not needed, their location is on the currently admit report\n')

# rewrite transplanted 18 mos export so that it's a real XL
print("import18mos")
month18filelist<- list.files(pattern = "18_M_TX", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
unique(rawepic$Organ)
rawepic <- rawepic %>% arrange(desc(`Txp Date`))
rawepic$`Txp Date` <- as.character(rawepic$`Txp Date`)
rawepic$`Organ Fail` <- as.character(rawepic$`Organ Fail`)
rawepic2 <- rawepic
rawepic2$Organ <- gsub("Heart\r\nKidney", "Heart,Kidney", rawepic$Organ)
rawepic2$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Pancreas", rawepic2$Organ)
unique(rawepic2$Organ)
rawepic3 <- rawepic2 %>%
	separate_longer_delim(c(Organ), delim = ",")
unique(rawepic3$Organ)
rawepichr <- subset(rawepic3, Organ == "Heart")
rawepickk <- subset(rawepic3, Organ == "Kidney")
rawepicli <- subset(rawepic3, Organ == "Liver")
rawepicpanc <- subset(rawepic3, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
print("import18mos2")
write_xlsx(rawepic3, "18mosraw.xlsx")
# load it and it works
wb <- loadWorkbook("18mosraw.xlsx")
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart 18 Months Transplanted","","","","","","","",""), rawepichr) # can comment 
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney 18 Months Transplanted","","","","","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver 18 Months Transplanted","","","","","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas 18 Months Transplanted","","","","","","","",""), rawepicpanc)
print("impot18mos3")
targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb,sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet))) # if comment title row above filter applied to row 1
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(33,10,10,11,8,20,8,11,45))
}
# for (sheet in sheet_names) {
  # # Read the sheet data
  # sheet_data <- read.xlsx(wb, sheet,startRow = 2)
  
  # # Find the column with "death" in its name
  # death_col <- grep("fail", tolower(colnames(sheet_data)), value = FALSE)
  
  # if (length(death_col) > 0) {
    # # If a "death" column is found, use its column letter in the Excel formula
    # death_col_letter <- LETTERS[death_col]
    
    # conditionalFormatting(wb, 
                          # sheet = sheet, 
                          # cols = 1:ncol(sheet_data), 
                          # rows = 3:2000, 
                          # rule = paste0("NOT(ISBLANK($", death_col_letter, "2))"), 
                          # style = negStyle)
    
    # print(paste("Conditional formatting applied to sheet:", sheet, "based on column:", colnames(sheet_data)[death_col]))
  # } else {
    # print(paste("No 'organ fail' column found in sheet:", sheet))
  # }
# }


saveWorkbook(wb, "18mosTxp History Print.xlsx", overwrite =TRUE)

# currently admitted
# rewrite export so that it's a real XL
month18filelist<- list.files(pattern = "Recipients_Currently_Admitted", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Pancreas", rawepic$Organ)
unique(rawepic$Organ)
# rawepic <- rawepic %>% arrange(desc(`Txp Date`))
rawepic2 <- rawepic %>%
	select(-c("DOB","Bed"))
rawepichr <- subset(rawepic2, Organ == "Heart")
rawepickk <- subset(rawepic2, Organ == "Kidney")
rawepicli <- subset(rawepic2, Organ == "Liver")
rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "CurrAdmitraw.xlsx")

# load it and it works
wb <- loadWorkbook("CurrAdmitraw.xlsx")
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart Currently Admitted","","","",""), rawepichr)
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Currently Admitted","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Currently Admitted","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Currently Admitted","","","",""), rawepicpanc)

targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(27,11,12,13,11))
}
saveWorkbook(wb, "CurrentlyAdmit Print.xlsx", overwrite =TRUE)

readline('\n "enter to open Tableau for reports"\n')
shell.exec(r'(M:\PAH\Transplant Quality\CMS\2022-2023\CMS Reports (8.18.2022).twb)')

readline('Copy crosstab Active WL then "enter to proceed"\n')
activeWL<-read_dataframe_from_clipboard()
readline('Copy crosstab rm due to death then "enter to proceed"\n')
rmduetodeath<-read_dataframe_from_clipboard()
readline('Copy crosstab rm due to other then "enter to proceed"\n')
rmduetoother<-read_dataframe_from_clipboard()
readline('Copy crosstab Eval not placed on list then "enter to proceed"\n')
evalnotlisted<-read_dataframe_from_clipboard()
readline('Copy crosstab LVD report then "enter to proceed"\n')
lvdreport<-read_dataframe_from_clipboard()

write_xlsx(activeWL, str_glue(r'({todaydir}\Active Waitlist Raw{todayfoldernm}.xlsx)'))
write_xlsx(rmduetodeath, str_glue(r'({todaydir}\Removed Due to Death Raw{todayfoldernm}.xlsx)'))
write_xlsx(rmduetoother, str_glue(r'({todaydir}\Removed Due to Other Raw{todayfoldernm}.xlsx)'))
write_xlsx(evalnotlisted, str_glue(r'({todaydir}\Eval Not Listed Raw{todayfoldernm}.xlsx)'))
write_xlsx(lvdreport, str_glue(r'({todaydir}\LVD Report Raw{todayfoldernm}.xlsx)'))

#tableaus
# Active WL
month18filelist<- list.files(pattern = "Active Waitlist Raw", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
colnames(rawepic) = gsub("Current.Phase.Of.Transplant.Status.NM", "WL.Status", colnames(rawepic))
colnames(rawepic) = gsub("Transplant.Episode.Organ.Name", "Organ", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Waitlist.DTS", "WL.DS", colnames(rawepic))
colnames(rawepic) = gsub("Patient.Age.When.Placed.On.Waitlist", "Age.on.WL", colnames(rawepic))
colnames(rawepic) = gsub("Medical.Record.NBR", "MRN", colnames(rawepic))
colnames(rawepic) = gsub("Patient.Race.NM", "Patient.Race", colnames(rawepic))
colnames(rawepic) = gsub("Patient.Sex.NM", "Gender", colnames(rawepic))
unique(rawepic$Organ)
# rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Panc", rawepic$Organ)

rawepic2 <- rawepic %>%
	select(-c("X","Center.Specific.Waitlist.DTS"))
rawepichr <- subset(rawepic2, Organ == "Heart")
rawepickk <- subset(rawepic2, Organ == "Kidney")
rawepicli <- subset(rawepic2, Organ == "Liver")
rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "Active Waitlist Print.xlsx")
month18filelist<- list.files(pattern = "Active Waitlist Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
wb <- loadWorkbook(month18filelist)
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart Active Waitlist","","","","","","",""), rawepichr)
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Active Waitlist","","","","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Active Waitlist","","","","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Active Waitlist","","","","","","",""), rawepicpanc)

targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(10,10,10,10,30,10,20,8))
}
saveWorkbook(wb, "Active Waitlist Print.xlsx", overwrite =TRUE)

# RM Due To Death
month18filelist<- list.files(pattern = "Removed Due to Death Raw", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
colnames(rawepic) = gsub("Transplant.Episode.Organ.Name", "Organ", colnames(rawepic))
colnames(rawepic) = gsub("Medical.Record.NBR", "MRN", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Waitlist.DTS", "WL.DS", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Waitlist.Removal.DS", "Rem.WL.DS", colnames(rawepic))
unique(rawepic$Organ)
# rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Panc", rawepic$Organ)

rawepic2 <- rawepic %>%
	select(-c("X","Current.Reason.For.Transplant.Phase.And.Status.NM"))
rawepichr <- subset(rawepic2, Organ == "Heart")
rawepickk <- subset(rawepic2, Organ == "Kidney")
rawepicli <- subset(rawepic2, Organ == "Liver")
rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "Removed Due to Death Print.xlsx")
month18filelist<- list.files(pattern = "Removed Due to Death Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
wb <- loadWorkbook(month18filelist)
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart Removed Due to Death","","","",""), rawepichr,
c("SPENCER,KENYARDA ALTWAND" ,"Heart" ,"905951977" ,"5/2/2022" ,"10/31/2022"),
c("SCOTT,CHARLES R JR." ,"Heart" ,"901139900" ,"10/11/2017" ,"12/01/2019"))
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Removed Due to Death","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Removed Due to Death","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Removed Due to Death","","","",""), rawepicpanc,
c("GOODMAN,CARLTON LEON" ,"Pancreas" ,"904335967" ,"12/13/2019" ,"08/28/2023"))


# rbind(c("Heart Removed Due to Death","","","",""),
# c("CRAWFORD,WILLIAM SEAN" ,"Heart" ,"904888256" ,"6/4/2024" ,"11/01/2024"),
# c("SPENCER,KENYARDA ALTWAND" ,"Heart" ,"905951977" ,"5/2/2022" ,"10/31/2022"),
# c("SCOTT,CHARLES R JR." ,"Heart" ,"901139900" ,"10/11/2017" ,"12/01/2019"),
# c("BOWMAN,TONY EUGENE" ,"Heart" ,"900804151" ,"10/19/2012" ,"10/10/2019"),
# c("STARR,JEREMY" ,"Heart" ,"903119167" ,"12/20/2018" ,"08/14/2019"),

# c("MEJIA,IVANOF" ,"Kidney/Pancreas" ,"904526031" ,"6/22/2023" ,"10/01/2024"),
# c("BECTON,DONNA SUZANNE" ,"Kidney/Pancreas" ,"908756284" ,"10/14/2022" ,"01/23/2024"),
# c("GOODMAN,CARLTON LEON" ,"Pancreas" ,"904335967" ,"12/13/2019" ,"08/28/2023"),
# c("FOX,CATHERINE SUZANNE" ,"Kidney/Pancreas" ,"905669889" ,"10/1/2021" ,"05/07/2022"),
# c("LIPHAM,HERBERT COREY" ,"Kidney/Pancreas" ,"901384530" ,"9/11/2018" ,"09/24/2020"),
# c("SMITH,CRYSTAL JENAE" ,"Kidney/Pancreas" ,"903987821" ,"5/9/2016" ,"03/25/2019"))



targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(30,15,11,11,11))
}
saveWorkbook(wb, "Removed Due to Death Print.xlsx", overwrite =TRUE)

#adding to the rm due to death rows


# RM Due To Other
month18filelist<- list.files(pattern = "Removed Due to Other Raw", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
rawepic$Waitlist.Removal.DS <- as.Date(rawepic$Waitlist.Removal.DS, format = "%m/%d/%Y",tz=Sys.timezone())
colnames(rawepic) = gsub("Transplant.Episode.Organ.Name", "Organ", colnames(rawepic))
colnames(rawepic) = gsub("Medical.Record.NBR", "MRN", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Waitlist.DTS", "WL.DS", colnames(rawepic))
colnames(rawepic) = gsub("Waitlist.Removal.DS", "Rem.WL.DS", colnames(rawepic))
colnames(rawepic) = gsub("Current.Reason.For.Transplant.Phase.And.Status.NM", "Reason", colnames(rawepic))

unique(rawepic$Organ)
# rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Panc", rawepic$Organ)
rawepic2 <- rawepic %>%
	select(-c("X","Day.of.Transplant.Surgery.DTS","WL.DS"))
rawepic2 <- rawepic2 %>% arrange(desc(Rem.WL.DS))
rawepic2$Rem.WL.DS <- as.character(rawepic2$Rem.WL.DS)
rawepic2$MRN <- as.character(rawepic2$MRN)
rawepichr <- subset(rawepic2, Organ == "Heart")
rawepickk <- subset(rawepic2, Organ == "Kidney")
rawepicli <- subset(rawepic2, Organ == "Liver")
rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "Removed Due to Other Print.xlsx")
month18filelist<- list.files(pattern = "Removed Due to Other Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
wb <- loadWorkbook(month18filelist)
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart Removed Due to Other","","","",""), rawepichr)
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Removed Due to Other","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Removed Due to Other","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Removed Due to Other","","","",""), rawepicpanc)

targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(30,15,11,11,25))
}
saveWorkbook(wb, "Removed Due to Other Print.xlsx", overwrite =TRUE)

# Evaled not listed
month18filelist<- list.files(pattern = "Eval Not Listed Raw", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
colnames(rawepic) = gsub("Transplant.Episode.Organ.Name", "Organ", colnames(rawepic))
colnames(rawepic) = gsub("Medical.Record.NBR", "MRN", colnames(rawepic))
colnames(rawepic) = gsub("Translant.Evaluation.End.DTS", "Eval.End", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Waitlist.Removal.DS", "Rem.WL.DS", colnames(rawepic))
colnames(rawepic) = gsub("Current.Reason.For.Transplant.Phase.And.Status.NM", "Reason", colnames(rawepic))

unique(rawepic$Organ)
# rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Panc", rawepic$Organ)
rawepic2 <- rawepic %>%
	select(-c("X","Decision.Date"))
rawepichr <- subset(rawepic2, Organ == "Heart")
rawepickk <- subset(rawepic2, Organ == "Kidney")
rawepicli <- subset(rawepic2, Organ == "Liver")
rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "Eval Not Listed Print.xlsx")
month18filelist<- list.files(pattern = "Eval Not Listed Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
wb <- loadWorkbook(month18filelist)
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart Evaled Not Listed","","","",""), rawepichr)
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Evaled Not Listed","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Evaled Not Listed","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Evaled Not Listed","","","",""), rawepicpanc)

targetorganscms<- copysubsetscms(rawepichr2,rawepickk2,rawepicli2,rawepicpanc2)
rawepichr2 <- targetorganscms[[1]]
rawepickk2 <- targetorganscms[[2]]
rawepicli2 <- targetorganscms[[3]]
rawepicpanc2 <- targetorganscms[[4]]

writeData(wb, hr, rawepichr2)
writeData(wb, kk, rawepickk2)
writeData(wb, li, rawepicli2)
writeData(wb, panc, rawepicpanc2)
removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(30,15,11,11,25))
}
saveWorkbook(wb, "Eval Not Listed Print.xlsx", overwrite =TRUE)

# LVD Report - we stopped doing liver living donors since 22'
month18filelist<- list.files(pattern = "LVD Report Raw", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
colnames(rawepic) = gsub("Transplant.Episode.Organ.Name", "Organ", colnames(rawepic))
colnames(rawepic) = gsub("Patient.MRN.ID", "MRN", colnames(rawepic))
colnames(rawepic) = gsub("Current.Phase.Of.Transplant.NM", "Phase", colnames(rawepic))
colnames(rawepic) = gsub("Min..Appointment.DTS", "Appt.DS", colnames(rawepic))
colnames(rawepic) = gsub("Current.Phase.Of.Transplant.Status.NM", "Status", colnames(rawepic))
colnames(rawepic) = gsub("Day.of.Donation.Date", "Donated", colnames(rawepic))
unique(rawepic$Organ)
# rawepic$Organ <- gsub("Kidney\r\nPancreas", "Kidney/Panc", rawepic$Organ)
rawepic2 <- rawepic %>%
	select(-c("Visit.Type.NM","Transplant.Episode.Record.ID","Min..Translant.Evaluation.Start.DS","X","Appointment.Status.NM"))
# rawepichr <- subset(rawepic2, Organ == "Heart")
# rawepickk <- subset(rawepic2, Organ == "Kidney")
# rawepicli <- subset(rawepic2, Organ == "Liver")
# rawepicpanc <- subset(rawepic2, Organ == "Pancreas" | Organ =="Kidney/Pancreas")
write_xlsx(rawepic2, "LVD Report Print.xlsx")
month18filelist<- list.files(pattern = "LVD Report Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
rawepic <- rbind(c("LVD Report","","","","","",""), rawepic)
wb <- loadWorkbook(month18filelist)
#if we ever go back to other organs for LVD
# hr <- addWorksheet(wb, sheetName = "Heart")
# kk <- addWorksheet(wb, sheetName = "Kidney")
# li <- addWorksheet(wb, sheetName = "Liver")
# panc <- addWorksheet(wb, sheetName = "Pancreas")
# writeData(wb, hr, rawepichr)
# writeData(wb, kk, rawepickk)
# writeData(wb, li, rawepicli)
# writeData(wb, panc, rawepicpanc)
# removeWorksheet(wb, "Sheet1")
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 1, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(38,11,11,12,11,14,11))
}
saveWorkbook(wb, "LVD Report Print.xlsx", overwrite =TRUE)
# move raw files away place DAR files in printme folder

dir.create(file.path(todaydirraw))
dir.create(file.path(todaydirprint))
rawfiles1 <- list.files(pattern = ".*raw.*\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
rawfiles2 <- list.files(pattern = "TXP_CMS", full.names = TRUE, ignore.case = TRUE)
darfiles <- list.files(pattern = "DAR", full.names = TRUE, ignore.case = TRUE)

suppressWarnings({
for (file in darfiles) {
  new_path <- normalizePath(file.path(todaydirprint, basename(file)), winslash = "/")
  old_path <- normalizePath(file, winslash = "/")
  
  if (file.exists(new_path)) {
    warning(paste("Destination file already exists:", new_path))
    next
  }
  
  success <- file.rename(from = old_path, to = new_path)
  if (!success) {
    warning(paste("Failed to rename:", old_path, "to", new_path))
  }
}

for (file in rawfiles1) {
  new_path <- normalizePath(file.path(todaydirraw, basename(file)), winslash = "/")
  old_path <- normalizePath(file, winslash = "/")
  
  if (file.exists(new_path)) {
    warning(paste("Destination file already exists:", new_path))
    next
  }
  
  success <- file.rename(from = old_path, to = new_path)
  if (!success) {
    warning(paste("Failed to rename:", old_path, "to", new_path))
  }
}

for (file in rawfiles2) {
  new_path <- normalizePath(file.path(todaydirraw, basename(file)), winslash = "/")
  old_path <- normalizePath(file, winslash = "/")
  
  if (file.exists(new_path)) {
    warning(paste("Destination file already exists:", new_path))
    next
  }
  
  success <- file.rename(from = old_path, to = new_path)
  if (!success) {
    warning(paste("Failed to rename:", old_path, "to", new_path))
  }
}

})
#make pdfs
cat('right click onthe pdf script and it will ask for directories you can paste from address bar and use the same dir twice if needed\n')
shell.exec('M:\\PAH\\Transplant Analyst Data\\duyXQ\\scripts\\outRealm\\ETL-R-main')

cat('CMS daily reports completed, great job\n')
}

TransplantRemovalinstructions <- function() {
cat('Login -> WL icon -> Removals \n enter to choose file \n')
shell.exec(r'(https://auth.unos.org/Login?redirect_uri=https%3A%2F%2Fportal.unos.org)')
exportedfilepath <- file.choose()
df <- read.csv(exportedfilepath, header = FALSE)

#does this change w/ seconds or not sometimes? pay attention to seps
df$removedtime <- as.POSIXct(df$V8,format="%Y-%m-%d %H:%M:%S",tz=Sys.timezone())

df415 <- subset(df, V9 == 4 | V9 == 15)

shell.exec(r'(M:\PAH\Transplant Analyst Data\Transplant Data Misc\TransplantAudits_Misc.twb)')
readline('Opening tabluea then toCK to clipboard then press enter\n')
tockWLrem <- read_dataframe_from_clipboard()
# filter out extra KPs maybe can do manually later..


#fix unos col names
colnames(df415)[colnames(df415) == "V1"] ="Patient.LNM"
colnames(df415)[colnames(df415) == "V2"] ="Patient.FNM"
colnames(df415)[colnames(df415) == "V3"] ="Patient.Social.Security.NBR"
colnames(df415)[colnames(df415) == "V4"] ="OrganNMUNOS"
colnames(df415)[colnames(df415) == "V5"] ="Age.At.Txp"
colnames(df415)[colnames(df415) == "V6"] ="Gender"
colnames(df415)[colnames(df415) == "V7"] ="ABO"
colnames(df415)[colnames(df415) == "V8"] ="RM.Date.UNOS"
colnames(df415)[colnames(df415) == "V9"] ="RM.Code"


# time convert with 12 hour format and am pm
tockWLrem$AnastomosisStartfix <- as.POSIXct(tockWLrem$Anastomosis.Start.DTS ,format="%m/%d/%Y %I:%M:%S %p",tz=Sys.timezone())

# ignore dupes via 000 or 999 ssns but true dupes, usually panc can be removed in post
colvec <- c("Patient.Social.Security.NBR")
getduperows(df415,colvec)
getduperows(tockWLrem,colvec)

# df415join <- df415 |> 
  # left_join(tockWLrem, by=c('Patient.Social.Security.NBR', 'Patient.LNM'='Patient.Last.NM'))

# df415join <- df415join %>% arrange(Transplant.Surgery.DTS)

# clipr::write_clip(df415join)

readline('There will be an error after this but its ignorable\n the data is on your clipboard, just paste into the most recent sheet in the folder "YEARWaitListRemoval"\nCheck Dates\n Look for duplicates - multi organ related\nKPs are generally ignored if both are under 24 hours.')
df415Rjoin <- df415 |> 
  right_join(tockWLrem, by=c('Patient.Social.Security.NBR'))
df415Rjoin$timetorm <- with(df415Rjoin, difftime(removedtime,AnastomosisStartfix,units="hours") )
df415Rjoin$Transplant.Surgery.DTS


df415Rjoin$Transplant.Surgery.DTS <- convertextdatemmddyyyy(df415Rjoin$Transplant.Surgery.DTS)

df415Rjoin <- df415Rjoin %>% arrange(Transplant.Surgery.DTS)
str(df415Rjoin)
df415Rjoin[, c('empty1', 'empty2', 'empty3')] = NA
outputdf415 <- df415Rjoin %>% select("Patient.Last.NM", "Patient.First.NM", "Patient.Birth.DTS", "Tx.Episode.Organ.UNOS.Definition", 
"AnastomosisStartfix", "Medical.Record.NBR",'empty1',"removedtime", 'empty2','empty3',"timetorm")

outputdf415 <- subset(outputdf415,!is.na(outputdf415$`timetorm`))
shell.exec(r'(M:\PAH\TX Audit\UNOS)')
clipr::write_clip(outputdf415, col.names = F, row.names = F)
cat('Paste into opened sheet\n')
}

AdverseEvents  <- function() {
readline("please start in 35w to press access buttons")
shell.exec(r'(M:\PAH\Transplant Analyst Data\RLs\AdverseEvents.twb)')
cat('Tableau is opening, change the dates to one week forward from current \n  -- If it gets behind, rule of thumb is 2 weeks before current for the more recent date and the older limit is ~6mos back \n \n')
#TX Detail
readline('Copy crosstab TX_Detail then "enter to proceed" \n')
rldetail<-read.delim("clipboard")
rldetail <- rldetail %>% select(-"X")
# str(rldetail)
wb <- loadWorkbook(r'(M:\PAH\Transplant Analyst Data\RLs\Archive\TXEpisodeDetail - Template.xlsx)')
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
}
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)
str(rldetail)
readline('setdiff below \n')
setdiff(names(rldetail),names(wbs))
readline('looking ok? Enter to delete old data and write into file \n')
# deleteData(wb, sheet = 1,  rows = 2:nrow(wbs)+10, cols = 1:ncol(wbs)+10, gridExpand = FALSE)
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)
 style <- 
        openxlsx::createStyle(
        wrapText = FALSE
       )
   # this adds the style to the excel file we just linked with R
# openxlsx::addStyle(wb, sheet = 1,  rows = 0:nrow(rldetail), cols = 0:ncol(rldetail), style, gridExpand = FALSE)
writeData(wb, sheet = "Sheet1", rldetail, colNames = F, rowNames = F, startRow = 2)
readline('Enter to write file \n')
saveWorkbook(wb,r'(M:\PAH\Transplant Analyst Data\RLs\TXEpisodeDetail.xlsx)',overwrite = T)


####RL DETAIL
readline('CHANGE DATE! \n Copy crosstab RL_Detail then "enter to proceed" \n')
{rldetail<-read.delim("clipboard")
rldetail <- rldetail %>% select(-"X")
rldetail$Event.Description.TXT <- scrubHTMLcodefromtext(rldetail$Event.Description.TXT)
str(rldetail)}
wb <- loadWorkbook(r'(M:\PAH\Transplant Analyst Data\RLs\Archive\RL_Detail - Template.xlsx)')
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
}
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)
str(rldetail)
readline('diff names coming up \n')
setdiff(names(rldetail),names(wbs))
readline('looking ok? Enter to delete old data and write into file \n')
# deleteData(wb, sheet = 1,  rows = 1:nrow(wbs)+10, cols = 0:ncol(wbs)+10, gridExpand = FALSE)
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)
# openxlsx::addStyle(wb, sheet = 1,  rows = 0:nrow(rldetail), cols = 0:ncol(rldetail), style, gridExpand = FALSE)
writeData(wb, sheet = "Sheet1", rldetail, colNames = F, rowNames = F, startRow = 2)
readline('Enter to write file \n')
saveWorkbook(wb,r'(M:\PAH\Transplant Analyst Data\RLs\RL_Detail.xlsx)',overwrite = T)


####RL DETAIL VAD
readline('Copy crosstab RL_Detail(VAD) then "enter to proceed" \n')
rldetail<-read.delim("clipboard")
rldetail <- rldetail %>% select(-"X")
rldetail$Event.Description.TXT <- scrubHTMLcodefromtext(rldetail$Event.Description.TXT)
str(rldetail)
wb <- loadWorkbook(r'(M:\PAH\Transplant Analyst Data\RLs\Archive\RL_Detail_VAD - Template.xlsx)')
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
}
wbs <- readWorkbook(wb, sheet = 1)
{str(wbs)
str(rldetail)
readline('below should just be "X"\n')
setdiff(names(rldetail),names(wbs))}
readline('looking ok? Enter to delete old data and write into file \n')
{# deleteData(wb, sheet = 1,  rows = 1:nrow(wbs)+10, cols = 0:ncol(wbs)+10, gridExpand = FALSE)
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)}
# openxlsx::addStyle(wb, sheet = 1,  rows = 0:nrow(rldetail)+1, cols = 0:ncol(rldetail)+1, style, gridExpand = FALSE)
writeData(wb, sheet = "Sheet1", rldetail, colNames = F, rowNames = F, startRow = 2)
readline('Enter to write file \n')
saveWorkbook(wb,r'(M:\PAH\Transplant Analyst Data\RLs\RL_Detail_VAD.xlsx)',overwrite = T)


####RL DETAIL ECMO
readline('Copy crosstab RL_Detail(ECMO) then "enter to proceed" \n')
# rldetail<-read.delim("clipboard")
# rldetail <- rldetail %>% select(-"X")
## clipr::write_clip(rldetail, row.names = F)
shell.exec(r'(M:\PAH\Transplant Analyst Data\RLs\RL_Detail_ECMO.xlsx)')
readline('opening XL to paste because scrubbing html breaks data somehow \n')
# rldetail$Event.Description.TXT <- scrubHTMLcodefromtext(rldetail$Event.Description.TXT)
# str(rldetail)
# wb <- loadWorkbook(r'(M:\PAH\Transplant Analyst Data\RLs\RL_Detail_ECMO - Template.xlsx)')
# sheet_names <- names(wb)
# for (sheet in sheet_names) {
# print(sheet)
# }
# wbs <- readWorkbook(wb, sheet = 1)
# {str(wbs)
# setdiff(names(rldetail),names(wbs))
# str(rldetail)
# readline('below should just be "X"\n')
# setdiff(names(rldetail),names(wbs))}
# readline('looking ok? Enter to delete old data and write into file \n')
# {
# deleteData(wb, sheet = 1,  rows = 1:nrow(wbs)+10, cols = 0:ncol(wbs)+10, gridExpand = FALSE)
# wbs <- readWorkbook(wb, sheet = 1)
# str(wbs)}
# {
# openxlsx::addStyle(wb, sheet = 1,  rows = 0:nrow(rldetail)+1, cols = 0:ncol(rldetail)+1, style, gridExpand = FALSE)
# writeData(wb, sheet = "Sheet1", rldetail, colNames = F, rowNames = F, startRow = 2)
# readline('Enter to write file \n')}
# saveWorkbook(wb,r'(M:\PAH\Transplant Analyst Data\RLs\RL_Detail_ECMO.xlsx)',overwrite = T)

###ECMOPatients
readline('do ECMOpts paste from tableau if theres an erorr because the times need to be formatted tried in code but should do manually still the times are working, but its the episode data\n')

ecmosheetpath <- r'(M:\PAH\CV\CVServices Admin\ECMO\ECMO Program Quality\ECMO Spreadsheets\PAH ECMO Program Patient List MASTER (20210707).xlsx)'
all_sheetsM <- importXLsheets(ecmosheetpath)
ecmosheetonly <- all_sheetsM$`Patient List`
rm(all_sheetsM) # why did i do this? i think its so i do the manual copy paste?
#possible col name change to watch
ecmosheetonly <- ecmosheetonly %>% select(c("Patient","MRN","Admit  date", "Date of Discharge","ECLS date","ECLS discontinued" ))
wb <- loadWorkbook(r'(M:\PAH\Transplant Analyst Data\RLs\Archive\ECMOPatients - Template.xlsx)')
sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
}
wbs <- readWorkbook(wb, sheet = 1)
{str(wbs)
setdiff(names(rldetail),names(wbs))
str(ecmosheetonly)

ecmosheetonly$`Admit  date`<- strftime( ecmosheetonly$`Admit  date`, format="%m/%d/%Y %k:%M", tz="UTC")
ecmosheetonly$`Date of Discharge`<- strftime( ecmosheetonly$`Date of Discharge`, format="%m/%d/%Y %k:%M", tz="UTC")
ecmosheetonly$`ECLS date`<- strftime( ecmosheetonly$`ECLS date`, format="%m/%d/%Y" , tz="UTC")
ecmosheetonly$`ECLS discontinued`<- strftime( ecmosheetonly$`ECLS discontinued`, format="%m/%d/%Y" , tz="UTC")
str(ecmosheetonly)

readline('below should just be "X"\n')
setdiff(names(ecmosheetonly),names(wbs))}
readline('looking ok? Enter to delete old data and write into file \n')
{
# deleteData(wb, sheet = 1,  rows = 1:nrow(wbs)+1, cols = 0:ncol(wbs)+1, gridExpand = FALSE)
wbs <- readWorkbook(wb, sheet = 1)
str(wbs)}
{
# openxlsx::addStyle(wb, sheet = 1,  rows = 0:nrow(ecmosheetonly)+1, cols = 0:ncol(ecmosheetonly)+1, style, gridExpand = FALSE)
writeData(wb, sheet = "Sheet1", ecmosheetonly, colNames = F, rowNames = F, startRow = 2)
readline('Enter to write file \n')
saveWorkbook(wb,r'(M:\PAH\Transplant Analyst Data\RLs\ECMOPatients.xlsx)',overwrite = T)
readline('paste ECMO patients because.. it doesnt work doing it direct from the file \n')}
shell.exec(r'(M:\PAH\Transplant Analyst Data\RLs\ECMOPatients.xlsx)')
shell.exec(r'(M:\PAH\Transplant Adverse Events\RiskEventTracking.accdb)')
readline('opening Access, import and run both "Final Append" queries then close access\n')
shell.exec(r'(M:\PAH\Transplant Adverse Events\RiskEventTrackingECMO.accdb)')
readline('opening Access for ECMO, load new RLs and use first query then Tableau will open\n')
readline('machine move required...\n')
clipr::write_clip("AdverseEvents73w()")
readline('go back to 73w for tableau driver function in clipboard\n ')
}
AdverseEvents73w  <- function() {
shell.exec(r'(M:\PAH\Transplant Adverse Events\Adverse Events_MM.twb)')
readline('opening Tableau, screenshot and double refresh detail view \n- observe graph changing\n')
readline('publish RL Review and save Tableau \n- enter will open next tableau VAD\n')

shell.exec(r'(M:\PAH\Transplant Adverse Events\AdverseEvents_VAD.twb)')
readline('Tableau, screenshot and double refresh detail view - observe graph update\n')
readline('publish AdverseEvents_VAD_DUY and save Tableau\n')

shell.exec(r'(M:\PAH\CV\CVServices Admin\ECMO\ECMO Adverse Events.twb)')
readline('screenshot and double refresh detail view - observe graph changning\n')
readline('publish ...ECMO_DUY and save Tableau\n')


}


AdverseEventsAccesstoTableau  <- function() {
readline('please be on 35w machine\n ')
shell.exec(r'(M:\PAH\Transplant Adverse Events\RiskEventTracking.accdb)')
readline('opening Access, import and run both "Final Append" queries then close access\n')
shell.exec(r'(M:\PAH\Transplant Adverse Events\RiskEventTrackingECMO.accdb)')
readline('opening Access for ECMO, load new RLs and use first query then Tableau will open\n')
readline('machine move required...\n')
clipr::write_clip("AdverseEvents73w()")
readline('go back to 73w for tableau driver function in clipboard\n ')

# shell.exec(r'(M:\PAH\Transplant Adverse Events\Adverse Events_MM.twb)')
# readline('opening Tableau, screenshot and double refresh detail view \n- observe graph changing\n')
# readline('publish RL Review and save Tableau \n- enter will open next tableau VAD\n')

# shell.exec(r'(M:\PAH\Transplant Adverse Events\AdverseEvents_VAD.twb)')
# readline('Tableau, screenshot and double refresh detail view - observe graph update\n')
# readline('publish AdverseEvents_VAD_DUY and save Tableau\n')

# shell.exec(r'(M:\PAH\CV\CVServices Admin\ECMO\ECMO Adverse Events.twb)')
# readline('screenshot and double refresh detail view - observe graph changning\n')
# readline('publish ...ECMO_DUY and save Tableau\n')


}

VadEcmoVolinstructions  <- function() {
shell.exec(r'(M:\PAH\Transplant Quality\4. Tableau\LVAD Tableau\Heart VAD ECMO Volume (20210824).twbx)')
clipr::write_clip(r'(M:\PAH\CV\CVServices Admin\ECMO\ECMO Program Quality\ECMO Spreadsheets\)')
readline('Starting Tableau \n Update ECMO DS to recent months volume \n publish without extract\n folder QA for now might move to Transplant\n filename _Duy \n')
readline('path is in your clipboard if DS disconnected -\n PAH ECMO Program Patient List MASTER (20210707).xlsx \n')
readline('ignore the second DS Ongoing...VADDyad.. seems it hasnt been updated in awhile \n')
readline('verify recent month is on there, now do the VAD volume \n')
clipr::write_clip(r'(M:\PAH\Transplant Quality\4. Tableau\)')
readline('path is in your clipboard if DS disconnected -\n VAD Outcomes Tracking Tool - BN.xlsx \n')
cat('Publish w/o extract then.... \n')
cat('Notify Stakeholders \n')
shell.exec('mailto:Daphne.Babin@piedmont.org;Kiersten.Ergen@piedmont.org;Kim.Munson@piedmont.org;Reannon.Wright@piedmont.org;Stephanie.Bass@piedmont.org;Zukari.Logan@piedmont.org?subject=Heart VAD ECMO Volume Tableau Updated&body=Monthly Heart VAD ECMO Volume updated.           https://tableau.piedmonthospital.org/#/workbooks/21557/views
')
cat('Thats all for this one you are done. \n')
}

MonthlyHeartCommittee  <- function() {
shell.exec(r'(M:\PAH\Transplant Analyst Data\Transplant Data Misc\TransplantAudits_Misc.twb)')
todayfoldernm <- readline('opening tableau \n')
todaydir<- Sys.Date()
# check if sub directory exists 
readline('Copy crosstab Heart Committee then "enter to proceed"')
activeWL<-read_dataframe_from_clipboard()
write_xlsx(activeWL, str_glue(r'(M:\PAH\Transplant Analyst Data\Reports\Heart Committee Audit Teandra\Heart Committee ran on {todaydir}.xlsx)'))
shell.exec(r'(M:\PAH\Transplant Analyst Data\Reports\Heart Committee Audit Teandra)')
cat('Email file to Teandra L. \n')

}
cusum <- function() {
shell.exec(r'(https://securesrtr.transplant.hrsa.gov/home/)')
}
AntiBodyBatchKits  <- function() {
readline('please use 35w \n \n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Dialysis_BatchOrders.twb)')
readline('Tableau is opening,  \n  -- copy crosstab for Transplant Episode and paste into listed pts excel sheet \n \n')
#TX Detail
shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\TransplantEpisodes_Listedpts.xlsx)')
readline('Copy crosstab DialysisHistory "enter to proceed" \n')
fixzips <- read_dataframe_from_clipboard()
fixzips$DialysisCenterZipCD <- sub("-.*", "", fixzips$DialysisCenterZipCD)
clipr::write_clip(fixzips)
shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\DialysisHistory.xlsx)')
readline('Replace all in sheet then \n Enter  \n \n \n')

source(r'(M:\PAH\Transplant Analyst Data\duyXQ\scripts\outRealm\ETL-R-main\batchkitformatting.r)')

# shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\DialysisOrders_EPIC _Current.accdb)')
# readline('Verify and Delete Dupes  \n  -- Close tab after deleting, then look at the other tab. \n \n')
# readline('Goto dropdown on left -> tables -> open listedpatientsall \n  -> Filter tubesto = blanks only, verify none are blank and close \n \n')
# readline('Run export_ShiptoDialysisCenter query, yes yes yes \n \n')
# readline('Open table "BatchOrderTemplate1" and sort by Client ID. \nEach record without a client will need to be looked up to see if it was assigned a Client ID. \n \n')
# readline('open batchordertemplate1\n \nlookup missing rows inside dialysis history XL or the old order \nalso use the globe dbo as a key \n-- enter will open current orders dir usually you can find places we already made\n')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Current Batch orders)')
# cat("We can join on Dialysis history to get rows\n")
# # extended antibody batch kits
# readline('copy the MRNs that are missing their centers\nfrom MS ACCESS and press enter \n')
# mrnt <- read_dataframe_from_clipboard()
# str(mrnt)
# readline('copy the full dialysis history we just pulled from tableau\nits also an excel sheet\n')
# dhist <- read_dataframe_from_clipboard()

# merged_df <- left_join(mrnt, dhist, 
                       # by = c("Patient.ID..MRN.EMPI." = "PatientMRNID"))


# clipr::write_clip(merged_df)
# cat("joined data is in clipboard, paste somewhere to manipulate it to fit into MS ACCESS \n\n")

# readline('export BatchOrderTemplate1 into that folder, these are the dialysis center kits\n \n')
# readline('query -> exportshipto home\n \n')
# readline('Clear rows in table -> batch order template _auxnew \n \n')
# readline('query -> "BatchOrderTemplate2 Without Matching BatchOrderTemplate_Home". This will add new home patients to the "batchOrderTemplate_AuxNew" table \n \n')
# readline('table -> batch order template _auxnew
# Update quantity column eg. if last month was 5 it will be 4 now \n- it will default to 6(only full orders) \n \n')
# readline('Export the table "BatchOrderTemplate_AuxNew" to \nan excel spreadsheet "BatchOrder_Homemmddyy" in the folder "M:\\PAH\\Transplant Analyst Data\\Antibody Kit Batch Orders\\Current Batch orders"\n \n')

# readline('queries -> Q_append aux batchordertemplateHOme \n \n')
# readline('save the master dialysis centers on close \n \n')
# readline('Check quantity columns before uploading \n \n')
# # get new excel and fix address from Joyce
# cat('replace DaVita Jonesboro Dialysis address with the right one 1595 Stockbridge Road\n\n PD1927 next\n')
shell.exec('https://www.spectrapath.net/')
readline('adding new center, search first, if not ther find newest PID\n add client and fill in\n Enter the name, address, phone number, dont update quantity')
readline('Dialysis file needs file prep on manual address changes, batchkitformatting.r')
readline('send ze email')
shell.exec('mailto:Evan.Matthews@piedmont.org;Joyce.Anderson@piedmont.org;Kristi.Goldberg@piedmont.org?subject=Antibody Batch Kits Uploaded&body=Uploaded and Attached Antibody Kit Batch Upload files.')
#new but only from jada guide
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders)')
# shell.exce(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\Dialysis_BatchOrders.twb)')
# readline("Copy Transplant Episode into sheet")
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\TransplantEpisodes_Listedpts.xlsx)')
# readline("Copy Dialysis History into sheet")
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Antibody Kit Batch Orders\DialysisHistory.xlsx)')
# readline("Open Access and Compile Data, yes to both imports")
# readline("\n\n\nClose import window after import complets")
# readline("\n\n\nclean up duplicates")
# readline("\n\n\n Open listed patients all table, check tubes to for Nulls")
# readline("\n\n\n Run Query Export_ShipToDialysisCtr")
# readline("\n\n\n Open table BatchOrderTemplate1, sort by client ID and fill in blanks
# client ID is found by cross ref dialysis history with master
# if it doesnt exist you need to add it to both places")
# readline("\n\n\n Spectrapath")
# shell.exec('https://www.spectrapath.net/')
# readline("\n\n\n Export the table Batch OrderTemplate1 to XL BatchOrder_Dialysismmddyy")
# readline("\n\n\n Run the Q_EmptyHomeTemplateCombinedTbl query \nto empty the HomeTemplate_Combined table.")
# readline("\n\n\n Empty records from table BatchOrderTemplate_AuxNew")
# readline("\n\n\n Run Export_ShipToHome")
# readline("\n\n\n Export the table BatchOrderTemplate2 XL BatchOrder_Homemmddyy")
# readline("\n\n\n Run Q_AppendAux_HomeTemplateCombinedTotalBatch")
# readline("\n\n\n Spectrapath, go to Orders tab, Batch Orders, BatchOrder_Dialysismmddyy")
# readline("\n\n\n Do same for BatchORderHome")
# readline("\n\n\n email sheets to Joyce Anderson, Evan Matthews, Angela Randall and all the pre kidney and liver coordinators")
# shell.exec('mailto:Evan.Matthews@piedmont.org,Joyce.Anderson@piedmont.org,Kristi.Goldberg@piedmont.org?subject=Antibody Batch Kits Uploaded')
}

Referralsmonthly  <- function() {

shell.exec(r'(C:\Users\246397\Desktop\Epic.lnk)')
readline('Login to Epic and run TX Referrals Stats_Monthly v3')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals)')
readline('Put in folder that is opened - filename is set by EPIC\n "enter to proceed" \n')
readline('Open file reclass NPI referred column to a number with 0 decimals and 
save as "Import_Spreadsheet" \n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\Mason Transplant Referrals.twb)')
readline('Openning Tableau....Copy crosstab "Providers_ForMasterTable"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\Tableau_ProviderRecords.xlsx)')
readline('Paste into this XL sheet save as "Tableau_ProviderRecords" in same folder, 
overwriting old excel file. 
\nChange data type in Provider ID column to TEXT!!!"\n')

readline('Enter to Open Access"\n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\ReferralsMasterDB FY23.accdb)')
readline('Go to Form1 and Click in "Import New Referrals" (import both sheets) This runs the macro "Mac_ImportNewReferrals"
\n')
readline('Open Table "Referrals_CurrentMonth" And sort by "referred by provider". 
Check to make sure all referrals by internal providers were deleted when the data was imported.
\n')
readline('Then Sort table by ZIP Field and fill in any missing counties 
(GA and AL residents only)here and on the provider record in EPIC \n')
readline('Then sort on "Type" column and delete any referrals for Ultrasound, Kidney Existing, 
or liver Existing\n')
readline('"Consultation". These will have to be looked up in EPIC and Changed to either 
"Liver Consultation" or "Kidney Consultation"\n')
readline('Save the Table \n')
readline('Click - Append New Referrals To "Referrals_All" Table \n')
readline('Click on the "Update referrals and Provider record for Tableau". This runs the macro "MacRefTbl"\n')
readline('This macro takes a few minutes.  Click Through the message pop ups.\n')
readline('duplicate table will popup check it Fr Fr\n')
readline('missing hub table will popup check it... kinda\n')
readline('edit SQL z_db_select_refferral months. add current month \n')
readline('save and run query \n')
readline('open table referral select months \n')
clipr::write_clip(r'(M:\PAH\Transplant Analyst Data\Referrals\Z_DB_ExportReferrals_SelectMonths.xlsx)')
readline('exportreferrals_selectmonths table\npath is in clipboard \n')
readline('filter for correct FYs (keeping only current and previous. Delete the rest\n')
readline('sort and color mrn cols to find dupes that are without a couple weeks of each other\n')
readline('paste select months contents into raw of referringNos file, they are both opening after you press enter \n')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\Z_DB_ExportReferrals_SelectMonths.xlsx)')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\DBabin Files\FY25\ReferringProvidersNos_July_FY25pubTEMPLATE.xlsx)')
readline('check for dupes in the id col in raw data, remove extra cols manually')
readline('reconnect data - click on pivot table and analyze -> reconnect check columns -> refresh')
readline('make sure fields cover the raw fields, refresh, cols will match two 
FYs if you did this right')
readline('updated raw data in DBabin folder, rename to Jul to curMonth and 
remove everything before TWO FYs in the raw data')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Referrals\DBabin Files\FY24\ReferringProvidersNos_July_Mar_FY24pubApr.xlsx)')
readline('email to daphne \n')
shell.exec('mailto:Daphne.Babin@piedmont.org;Teandra.Lassiter@piedmont.org;Jennifer.Wilcox@piedmont.org?subject=Referral Volumes&body=Attached.
')
readline('Tableau\n\n On the summary Dashboard, Change the "Referrals through: date\n')
readline('Go to datasource and refresh the data source to reflect the new referrals and then republish workbook. It wont repub..\n')


# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyListedKDPI.xlsx)')



# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\DataSourceForCPRAMELDScores.twb)')
# cat('starting tableau for copy crosstab procedures\n')
# readline('Tableau should be up \n 1. "enter to proceed"\n')
# readline('We are going to copy paste each sheet into a new file that will open \n "enter to proceed"\n')

# readline('Copy to Crosstab "Dialysis Type" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList_DialyisType.xlsx)')
# readline('Paste under first row and check headers then delete new headers row  \n Reclass: as text MRN, SSN, ORGAN\n')

# readline('Copy to Crosstab "FullEvalProtocols" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyFullEval_Protocol.xlsx)')
# readline('Paste at end of file - delete overlap this will have only new appts  \n Reclass: as text MRN, SSN, ORGAN\n')

# readline('Copy to Crosstab "KidneyWaitlist" Tableau Worksheet and opening EXCEL \n "enter to proceed"\n')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyTransplantEpisodeData.xlsx)')
# readline('Paste under first row and check headers then delete new headers row  \n Reclass: as text MRN, SSN, ORGAN\n')

# cat('moving old files \n KidneywaitlistDataFullEvals_All, \n QData_ForTop50_ThursdayRept, \n QData_KDPI_ThurdaysRept')
# shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneywaitlistDataFullEvals_All.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
# shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\QData_ForTop50_ThursdayRept.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
# shell(r'(move "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\QData_KDPI_ThurdaysRept.xlsx" "M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Archive")')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\KidneyWaitList.accdb)')
# cat('starting access')
# readline('export 2 QDATA tables into the Kidney WL Report dir \n QDATA = KidneywaitlistDataFullEvals_All for file name')
# readline('then run 2 Qdata queries that will make 2 more Qdata tables to export')
# readline('open QDATA KDPI and paste into first sheet excel to open')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\Thursday Report.xlsx)')
# shell.exec(r'(M:\PAH\Transplant Analyst Data\Kidney Wait List Report\CPRAs_MELDScores.twb)')
# readline('Keeping thursday Report open paste pink sheet into second Excel sheet \n should come from Weekly full eval active tab in Tableau')
# readline('Save and email to Kristi :] you did it yay!')


readline('Task Complete remove from calendar for tracking  \n "enter to proceed"\n')
# return(a2b2env)
}

#path error if you have the file open.
yellowStyle <- createStyle(fontColour = "#000000", bgFill = "#FFFF00")
negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
posStyle <- createStyle(fontColour = "#000000", bgFill = "#D1FFBD")
CLIPR_ALLOW=TRUE
print(Sys.setenv(CLIPR_ALLOW = TRUE))  # `A+C` could also be used
Sys.getenv("CLIPR_ALLOW")
# getduperowslearntreturnmultiple <- function(datafrm,colvec){
# dupfromcolvec <- sort(datafrm[[colvec]][duplicated(datafrm[[colvec]]) | duplicated(datafrm[[colvec]], fromLast=TRUE)])
# duperows <- subset(datafrm,datafrm[[colvec]] %in% dupfromcolvec) %>% arrange(colvec)
# return(c(dupfromcolvec, duperows))
# single assignment will get all items in the list returned