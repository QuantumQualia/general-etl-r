#Remaining at risk
# this needs an edit... important cohort is the one with any still left so the 3mos ends first so we need the list from like the third cohort? idk the goal is to put a valuable list on qapi slides.

source("M:\\PAH\\Transplant Analyst Data\\duyXQ\\scripts\\outRealm\\ETL-R-main\\sourcery.R")
today <- Sys.Date() # duplicated... will rn later
readline('Export SCRallorgans from EPIC path in clipboard and opening the folder just incase\n')
#shell.exec(r'(C:\Users\246397\Desktop\Epic.lnk)')
shell.exec(r'(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\archive)')
clipr::write_clip(r'(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\archive)')
readline('Enter to begin importing the XL wait for report to FULLY Load\n')
QRattendfiles<- list.files(path = r"(M:\PAH\Transplant Analyst Data\Reports\multiOrganReports\archive\)" ,pattern = "sCrallorgans", full.names = TRUE, ignore.case = TRUE)
file_info <- file.info(QRattendfiles)
file_info$file_name <- rownames(file_info)
sorted_files <- QRattendfiles[order(desc(file_info$ctime))]
qrAttend<-sorted_files[1]
# print(sorted_files)
#joining attendance sheets
qrAttend2 <- importXLsheets(qrAttend)
rawepic <- as.data.frame(read_excel(qrAttend))
str(rawepic)


rawepic <- rawepic %>%
  mutate(followupend1yr = `Txp Date` + 365*60*60*24,
		followupend90d = `Txp Date` + 90*60*60*24) # adding a year and 90 days

#subset cohorts by followup Date
recips2 <- subset(rawepic,`Donor/Recip` == "Recipient" & `Is Multi-organ` == "N" & !is.na(`Admission Date`))
str(recips2)
# recips <- subset(recips, as.Date(followupend1yr) >= Sys.Date())
recips2$organfaildate <- coalesce(recips2$`Organ Fail...27` ,recips2$`Organ Fail...28`)
colnames(recips2)[colnames(recips2) == "Episode ID"] ="EpisodeID"

recips3 <- recips2 %>% select(c("EpisodeID","MRN","Patient Name", "DOB","Age at Txp","Admission Date","Center Waitlisted","Txp Date","followupend90d", "followupend1yr", "Last CR", "Warm Ischemia Time (min)","Total Ischemia Time (min)","Death Date","organfaildate","Organ"))
str(recips3)
recips4 <-  recips3 %>% arrange(followupend1yr)
recips5 <- recips4 %>% 
	mutate_if(is.POSIXct, as.Date) #ispos in sourcery

df <- recips5

# Ensure date columns are in Date format
df$`Txp Date` <- as.Date(df$`Txp Date`)
df$organfaildate <- as.Date(df$organfaildate)
df$followupend90d <- as.Date(df$followupend90d)

today_date <- Sys.Date()

df$highlight_flag <- with(df, ifelse(
  !is.na(organfaildate) & !is.na(`Txp Date`) & abs(organfaildate - `Txp Date`) <= 365, 3,  # 
  ifelse(!is.na(followupend90d) & followupend90d < today_date & is.na(organfaildate),1,   # 
         ifelse(!is.na(organfaildate), 2, NA)  
  )
))

df_sorted <- df[order(-df$highlight_flag), ]

df_sorted$highlight_flag <- NULL

recips <- df_sorted


# shift important columnup


# cohort config 
# change years, shift dates in pairs
# just change dates if really needed, titles can change in the file
jul24d1 <- as.Date("2022-07-01")
jul24d2 <- as.Date("2024-12-31")
jul24cutoff90d <- jul24d2 + 90
jul24cutoff1yr <- jul24d2 + 182
jan25d1 <- as.Date("2023-01-01")
jan25d2 <- as.Date("2025-06-30")
jan25cutoff90d <- jan25d2 + 90
jan25cutoff1yr <- jan25d2 + 182
jul25d1 <- as.Date("2023-07-01")
jul25d2 <- as.Date("2025-12-31")
jul25cutoff90d <- jul25d2 + 90
jul25cutoff1yr <- jul25d2 + 182
jan26d1 <- as.Date("2024-01-01")
jan26d2 <- as.Date("2026-06-30")
jan26cutoff90d <- jan26d2 + 90
jan26cutoff1yr <- jan26d2 + 182

# txp date subset for each CH & limit follow up times for dates that extend past cutoff
Jul2024CH <- subset(recips,as.Date(`Txp Date`) >= jul24d1 & as.Date(`Txp Date`) <= jul24d2)
Jul2024CH$followupend90d[Jul2024CH$followupend90d >= jul24cutoff90d] <- jul24cutoff90d
Jul2024CH$followupend1yr[Jul2024CH$followupend1yr >= jul24cutoff1yr] <- jul24cutoff1yr
Jan2025CH <- subset(recips,as.Date(`Txp Date`) >= jan25d1 & as.Date(`Txp Date`) <= jan25d2)
Jan2025CH$followupend90d[Jan2025CH$followupend90d >= jan25cutoff90d] <- jan25cutoff90d
Jan2025CH$followupend1yr[Jan2025CH$followupend1yr >= jan25cutoff1yr] <- jan25cutoff1yr
Jul2025CH <- subset(recips,as.Date(`Txp Date`) >= jul25d1 & as.Date(`Txp Date`) <= jul25d2)
Jul2025CH$followupend90d[Jul2025CH$followupend90d >= jul25cutoff90d] <- jul25cutoff90d
Jul2025CH$followupend1yr[Jul2025CH$followupend1yr >= jul25cutoff1yr] <- jul25cutoff1yr
Jan2026CH <- subset(recips,as.Date(`Txp Date`) >= jan26d1 & as.Date(`Txp Date`) <= jan26d2)
Jan2026CH$followupend90d[Jan2026CH$followupend90d >= jan26cutoff90d] <- jan26cutoff90d
Jan2026CH$followupend1yr[Jan2026CH$followupend1yr >= jan26cutoff1yr] <- jan26cutoff1yr

# filter out those past cutoffs today
#Jul2024CH90datrisk <- subset(Jul2024CH,as.Date(followupend90d) >= Sys.Date() | `Death Date` <= as.Date(followupend90d) | organfaildate <= as.Date(followupend90d))
#Jul2024CH1YRatrisk <- subset(Jul2024CH,as.Date(followupend1yr) >= Sys.Date() | `Death Date` <= as.Date(followupend1yr) | organfaildate <= as.Date(followupend1yr))

#Jan2025CH90datrisk <- subset(Jan2025CH,as.Date(followupend90d) >= Sys.Date() | `Death Date` <= as.Date(followupend90d) | organfaildate <= as.Date(followupend90d))
#Jan2025CH1YRatrisk <- subset(Jan2025CH,as.Date(followupend1yr) >= Sys.Date() | `Death Date` <= as.Date(followupend1yr) | organfaildate <= as.Date(followupend1yr))

#Jul2025CH90datrisk <- subset(Jul2025CH,as.Date(followupend90d) >= Sys.Date() | `Death Date` <= as.Date(followupend90d) | organfaildate <= as.Date(followupend90d))
#Jul2025CH1YRatrisk <- subset(Jul2025CH,as.Date(followupend1yr) >= Sys.Date() | `Death Date` <= as.Date(followupend1yr) | organfaildate <= as.Date(followupend1yr))

#Jan2026CH90datrisk <- subset(Jan2026CH,as.Date(followupend90d) >= Sys.Date() | `Death Date` <= as.Date(followupend90d) | organfaildate <= as.Date(followupend90d))
#Jan2026CH1YRatrisk <- subset(Jan2026CH,as.Date(followupend1yr) >= Sys.Date() | `Death Date` <= as.Date(followupend1yr) | organfaildate <= as.Date(followupend1yr))

#without removing no longer at risk folks

Jul2024CH90datrisk <- Jul2024CH
Jul2024CH1YRatrisk <- Jul2024CH

Jan2025CH90datrisk <- Jan2025CH
Jan2025CH1YRatrisk <- Jan2025CH

Jul2025CH90datrisk <- Jul2025CH
Jul2025CH1YRatrisk <- Jul2025CH

Jan2026CH90datrisk <- Jan2026CH
Jan2026CH1YRatrisk <- Jan2026CH


Jul2024CH90datrisk <-Jul2024CH90datrisk %>% select(-c(followupend1yr))
Jul2024CH1YRatrisk <-Jul2024CH1YRatrisk %>% select(-c(followupend90d))
				
Jan2025CH90datrisk <-Jan2025CH90datrisk %>% select(-c(followupend1yr))
Jan2025CH1YRatrisk <-Jan2025CH1YRatrisk %>% select(-c(followupend90d))
				
Jul2025CH90datrisk <-Jul2025CH90datrisk %>% select(-c(followupend1yr))
Jul2025CH1YRatrisk <-Jul2025CH1YRatrisk %>% select(-c(followupend90d))

Jan2026CH90datrisk <-Jan2026CH90datrisk %>% select(-c(followupend1yr))
Jan2026CH1YRatrisk <-Jan2026CH1YRatrisk %>% select(-c(followupend90d))

#cheque
nrow(Jul2024CH90datrisk) # empty
nrow(Jul2024CH1YRatrisk)
nrow(Jan2025CH90datrisk)
nrow(Jan2025CH1YRatrisk)
nrow(Jul2025CH90datrisk)
nrow(Jul2025CH1YRatrisk)
nrow(Jan2026CH90datrisk)
nrow(Jan2026CH1YRatrisk)

#organ subsets
Jul2024CH90datriskKid <- subset(Jul2024CH90datrisk, Organ == "Kidney")
Jul2024CH1YRatriskKid <- subset(Jul2024CH1YRatrisk, Organ == "Kidney")
Jan2025CH90datriskKid <- subset(Jan2025CH90datrisk, Organ == "Kidney")
Jan2025CH1YRatriskKid <- subset(Jan2025CH1YRatrisk, Organ == "Kidney")
Jul2025CH90datriskKid <- subset(Jul2025CH90datrisk, Organ == "Kidney")
Jul2025CH1YRatriskKid <- subset(Jul2025CH1YRatrisk, Organ == "Kidney")
Jan2026CH90datriskKid <- subset(Jan2026CH90datrisk, Organ == "Kidney")
Jan2026CH1YRatriskKid <- subset(Jan2026CH1YRatrisk, Organ == "Kidney")


#with recent CH OR
sheets <- list(
"Jan2026-90datrisk" = Jul2024CH90datriskKid,
"Jan2026-1YRatrisk" = Jul2024CH1YRatriskKid,
"Jul2026-90datrisk" = Jan2025CH90datriskKid,
"Jul2026-1YRatrisk" = Jan2025CH1YRatriskKid, 
"Jan2027-90datrisk" = Jul2025CH90datriskKid, 
"Jan2027-1YRatrisk" = Jul2025CH1YRatriskKid,
"Jul2027-90Datrisk" = Jan2026CH90datriskKid,
"Jul2027-1YRatrisk" = Jan2026CH1YRatriskKid) #assume sheet1 and sheet2 are data frames
#####without recents
# sheets <- list("Jul2026-90datrisk" = Jan2025CH90datriskKid,"Jul2026-1YRatrisk" = Jan2025CH1YRatriskKid, "Jan2027-90datrisk" = Jul2025CH90datriskKid, "Jan2027-1YRatrisk" = Jul2025CH1YRatriskKid)

write_xlsx(sheets, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\KidneyRemainingAtRisk2.xlsx")
wb <- loadWorkbook("M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\KidneyRemainingAtRisk2.xlsx")
#sheet_names <- names(wb)
#for (sheet in sheet_names) {
#print(sheet)
#freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
#addFilter(wb, sheet =sheet, rows = 1, cols =1:ncol(read.xlsx(wb,sheet)))
#conditionalFormatting(wb, sheet=sheet, cols =1:ncol(read.xlsx(wb,sheet)), rows=2:2000, rule="NOT(ISBLANK($N2))", style = negStyle)
#setColWidths(wb, sheet= sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(15,15,25,15,15,15,15,15,15,15,15,15,15,15,15))
#}
sheet_names <- names(wb)
for (sheet in sheet_names) {
  print(sheet)
  freezePane(wb, sheet = sheet, firstRow = TRUE, firstCol = FALSE)
  addFilter(wb, sheet = sheet, rows = 1, cols = 1:ncol(read.xlsx(wb, sheet)))
  
  conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "NOT(ISBLANK($N2))", 
    style = yellowStyle
  )

  conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "AND(ISNUMBER($N2), ISNUMBER($H2), ABS($N2 - $H2) <= 365)", 
    style = negStyle
  )
  
  conditionalFormatting(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
                        rows = 2:2000, 
                        rule = paste0("AND(NOT(ISBLANK($I2)),$I2<DATE(", year(today), ",", month(today), ",", day(today), "),ISBLANK($N2))"), 
                        style = posStyle)
  
  setColWidths(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
               widths = c(15,15,25,15,15,15,15,15,15,15,15,15,15,15,15))
}
saveWorkbook(wb, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\KidneyRemainingAtRiskformat.xlsx", overwrite =TRUE)





Jul2024CH90datriskLiv <- subset(Jul2024CH90datrisk, Organ == "Liver")
Jul2024CH1YRatriskLiv <- subset(Jul2024CH1YRatrisk, Organ == "Liver")
Jan2025CH90datriskLiv <- subset(Jan2025CH90datrisk, Organ == "Liver")
Jan2025CH1YRatriskLiv <- subset(Jan2025CH1YRatrisk, Organ == "Liver")
Jul2025CH90datriskLiv <- subset(Jul2025CH90datrisk, Organ == "Liver")
Jul2025CH1YRatriskLiv <- subset(Jul2025CH1YRatrisk, Organ == "Liver")
Jan2026CH90datriskLiv <- subset(Jan2026CH90datrisk, Organ == "Liver")
Jan2026CH1YRatriskLiv <- subset(Jan2026CH1YRatrisk, Organ == "Liver")

#with recent CH OR
sheets <- list(
"Jan2026-90datrisk" = Jul2024CH90datriskLiv, 
"Jan2026-1YRatrisk" = Jul2024CH1YRatriskLiv, 
"Jul2026-90datrisk" = Jan2025CH90datriskLiv,
"Jul2026-1YRatrisk" = Jan2025CH1YRatriskLiv, 
"Jan2027-90datrisk" = Jul2025CH90datriskLiv, 
"Jan2027-1YRatrisk" = Jul2025CH1YRatriskLiv,
"Jul2027-90Datrisk" = Jan2026CH90datriskLiv,
"Jul2027-1YRatrisk" = Jan2026CH1YRatriskLiv)
#####without recents
# sheets <- list("Jul2026-90datrisk" = Jan2025CH90datriskLiv,"Jul2026-1YRatrisk" = Jan2025CH1YRatriskLiv, "Jan2027-90datrisk" = Jul2025CH90datriskLiv, "Jan2027-1YRatrisk" = Jul2025CH1YRatriskLiv)
write_xlsx(sheets, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\LiverRemainingAtRisk2.xlsx")

wb <- loadWorkbook("M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\LiverRemainingAtRisk2.xlsx")
sheet_names <- names(wb)
for (sheet in sheet_names) {
  print(sheet)
  freezePane(wb, sheet = sheet, firstRow = TRUE, firstCol = FALSE)
  addFilter(wb, sheet = sheet, rows = 1, cols = 1:ncol(read.xlsx(wb, sheet)))
  
conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "NOT(ISBLANK($N2))", 
    style = yellowStyle
  )

  conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "AND(ISNUMBER($N2), ISNUMBER($H2), ABS($N2 - $H2) <= 365)", 
    style = negStyle
  )
  conditionalFormatting(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
                        rows = 2:2000, 
                        rule = paste0("AND(NOT(ISBLANK($I2)),$I2<DATE(", year(today), ",", month(today), ",", day(today), "),ISBLANK($N2))"), 
                        style = posStyle)
  
  setColWidths(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
               widths = c(15,15,25,15,15,15,15,15,15,15,15,15,15,15,15))
}
saveWorkbook(wb, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\LiverRemainingAtRiskformat.xlsx", overwrite =TRUE)

Jul2024CH90datriskHtx <- subset(Jul2024CH90datrisk, Organ == "Heart")
Jul2024CH1YRatriskHtx <- subset(Jul2024CH1YRatrisk, Organ == "Heart")
Jan2025CH90datriskHtx <- subset(Jan2025CH90datrisk, Organ == "Heart")
Jan2025CH1YRatriskHtx <- subset(Jan2025CH1YRatrisk, Organ == "Heart")
Jul2025CH90datriskHtx <- subset(Jul2025CH90datrisk, Organ == "Heart")
Jul2025CH1YRatriskHtx <- subset(Jul2025CH1YRatrisk, Organ == "Heart")
Jan2026CH90datriskHtx <- subset(Jan2026CH90datrisk, Organ == "Heart")
Jan2026CH1YRatriskHtx <- subset(Jan2026CH1YRatrisk, Organ == "Heart")

#with recent CH OR replace Jan2020- with next and previous
sheets <- list(
"Jan2026-90datrisk" = Jul2024CH90datriskHtx,
"Jan2026-1YRatrisk" = Jul2024CH1YRatriskHtx,
"Jul2026-90datrisk" = Jan2025CH90datriskHtx,
"Jul2026-1YRatrisk" = Jan2025CH1YRatriskHtx, 
"Jan2027-90datrisk" = Jul2025CH90datriskHtx, 
"Jan2027-1YRatrisk" = Jul2025CH1YRatriskHtx,
"Jul2027-90Datrisk" = Jan2026CH90datriskHtx,
"Jul2027-1YRatrisk" = Jan2026CH1YRatriskHtx)
#####without recents
# sheets <- list("Jul2026-90datrisk" = Jan2025CH90datriskHtx,"Jul2026-1YRatrisk" = Jan2025CH1YRatriskHtx, "Jan2027-90datrisk" = Jul2025CH90datriskHtx, "Jan2027-1YRatrisk" = Jul2025CH1YRatriskHtx)

write_xlsx(sheets, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\HeartRemainingAtRisk2.xlsx")
wb <- loadWorkbook("M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\HeartRemainingAtRisk2.xlsx")
sheet_names <- names(wb)
for (sheet in sheet_names) {
  print(sheet)
  freezePane(wb, sheet = sheet, firstRow = TRUE, firstCol = FALSE)
  addFilter(wb, sheet = sheet, rows = 1, cols = 1:ncol(read.xlsx(wb, sheet)))
  
conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "NOT(ISBLANK($N2))", 
    style = yellowStyle
  )

  conditionalFormatting(
    wb, 
    sheet = sheet, 
    cols = 1:ncol(read.xlsx(wb, sheet)), 
    rows = 2:2000, 
    rule = "AND(ISNUMBER($N2), ISNUMBER($H2), ABS($N2 - $H2) <= 365)", 
    style = negStyle
  )
  conditionalFormatting(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
                        rows = 2:2000, 
                        rule = paste0("AND(NOT(ISBLANK($I2)),$I2<DATE(", year(today), ",", month(today), ",", day(today), "),ISBLANK($N2))"), 
                        style = posStyle)
  
  setColWidths(wb, sheet = sheet, cols = 1:ncol(read.xlsx(wb, sheet)), 
               widths = c(15,15,25,15,15,15,15,15,15,15,15,15,15,15,15))
}
saveWorkbook(wb, "M:\\PAH\\Transplant Analyst Data\\Reports\\multiOrganReports\\HeartRemainingAtRiskformat.xlsx", overwrite =TRUE)
readline("sort by yellow then red and remove closed out cohorts file is created\n")
toddy <- Sys.Date()
shell.exec(str_glue(r'(mailto:Clark.Kensinger@piedmont.org;Sundus.Lodhi@piedmont.org;Kristi.Goldberg@piedmont.org?subject=SRTR Kidney Remaining At Risk 90D 1YR [{toddy}]&cc=Teandra.Lassiter@piedmont.org;Desriee.Plummer@piedmont.org&body=File Attached. %0D%0A%0D%0A
Background color key:%0D%0A
Red = flagged within 365 days of transplant%0D%0A
Yellow = flagged > 365 days of transplant%0D%0A
Green = follow-up ended not flagged%0D%0A
Colorless = follow-up hasn’t ended %0D%0A
)'))
shell.exec(str_glue(r'(mailto:Lance.Stein@piedmont.org;Elizabeth.Miller1@piedmont.org;Jonathan.Hundley@piedmont.org?subject=SRTR Liver Remaining At Risk 90D 1YR [{toddy}]&cc=Teandra.Lassiter@piedmont.org;Desriee.Plummer@piedmont.org&body=File Attached. %0D%0A%0D%0A
Background color key:%0D%0A
Red = flagged within 365 days of transplant%0D%0A
Yellow = flagged > 365 days of transplant%0D%0A
Green = follow-up ended not flagged%0D%0A
Colorless = follow-up hasn’t ended %0D%0A)'))
shell.exec(str_glue(r'(mailto:Ezequiel.Molina@piedmont.org;Elena.Gozlan@piedmont.org?subject=SRTR Heart Remaining At Risk 90D 1YR [{toddy}]&cc=Teandra.Lassiter@piedmont.org;Desriee.Plummer@piedmont.org;Taly.Moua@piedmont.org&body=File Attached. %0D%0A%0D%0A
Background color key:%0D%0A
Red = flagged within 365 days of transplant%0D%0A
Yellow = flagged > 365 days of transplant%0D%0A
Green = follow-up ended not flagged%0D%0A
Colorless = follow-up hasn’t ended %0D%0A)'))
#accomplished above
# Jan2025CHrem <- subset(Jan2025CH, as.Date(followupend) >= Sys.Date())
# Jul2025CHrem <- subset(Jul2025CH, as.Date(followupend) >= Sys.Date())

# ----
# maturity <- Sys.Date()
# n  <- 6
# bytime <- paste("-",n," months",sep="")
# sixMosago <- seq(maturity,length.out=2,by=bytime)[2]