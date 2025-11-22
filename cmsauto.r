source(r'(M:\PAH\Transplant Analyst Data\Analyst Documentation\Automated\sourcefiles\scriptbank.r)')

userid <- readline('enter your userid\n')

todayfoldernm <- readline('enter dir for reports eg 03.20.2024: \n')
todaydir<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm})')
todaydirraw<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm}\raw)')
todaydirprint<- str_glue(r'(M:\PAH\Transplant Quality\CMS\2022-2023\Daily Reports\FY 2025\{todayfoldernm}\printme)')
# check if sub directory exists 
if (dir.exists(todaydir)){
	cat('todaydir exists')
	shell.exec(todaydir)
} else {
	dir.create(file.path(todaydir))
	shell.exec(todaydir)
}
#shell.exec(str_glue(r'(C:\Users\{userid}\Desktop\Epic.lnk)'))
clipr::write_clip(todaydir)
setwd(todaydir)
print(todaydir)
readline('Login to Epic and run reports \n 1. CMS Currently Admitted \n 2. Transplanted 18 months \n\n Export both into dir that is opened\npath already in clipboard\n')
readline('3. Print the DAR - one for MTP and one for Heart Failure Center \n- Report is made but if needed to build...\n Change visit type to List Post Tx follow, living donor follow\n')
readline('Be sure to name them DAR Heart Print and DAR Abdominal Print\n')
# rewrite transplanted 18 mos export so that it's a real XL
cat('checking inputs...\n')
month18filelist<- list.files(pattern = "18_M_TX", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
unique(rawepic$Organ)
rawepic <- rawepic %>% arrange(desc(`Txp Date`))
rawepic$`Txp Date` <- as.character(rawepic$`Txp Date`)
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

write_xlsx(rawepic3, "18mosraw.xlsx")
wb <- loadWorkbook("18mosraw.xlsx")
hr <- addWorksheet(wb, sheetName = "Heart")
rawepichr2 <- rbind(c("Heart 18 Months Transplanted","","","",""), rawepichr) 
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney 18 Months Transplanted","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver 18 Months Transplanted","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas 18 Months Transplanted","","","",""), rawepicpanc)

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
addFilter(wb, sheet =sheet, rows = 2, cols =1:ncol(read.xlsx(wb,sheet))) # if comment title row above filter applied to row 1
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(33,10,10,11,11))
}
saveWorkbook(wb, "18mosTxp History Print.xlsx", overwrite =TRUE)
# currently admitted
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
cat(str_glue('Number of rows in data imported: {nrow(activeWL)}'), '\n'))
readline('Copy crosstab rm due to death then "enter to proceed"\n')
rmduetodeath<-read_dataframe_from_clipboard()
cat(str_glue('Number of rows in data imported: {nrow(rmduetodeath)}'), '\n'))
readline('Copy crosstab rm due to other then "enter to proceed"\n')
rmduetoother<-read_dataframe_from_clipboard()
cat(str_glue('Number of rows in data imported: {nrow(rmduetoother)}'), '\n'))
readline('Copy crosstab Eval not placed on list then "enter to proceed"\n')
evalnotlisted<-read_dataframe_from_clipboard()
cat(str_glue('Number of rows in data imported: {nrow(evalnotlisted)}'), '\n'))
readline('Copy crosstab LVD report then "enter to proceed"\n')
lvdreport<-read_dataframe_from_clipboard()
cat(str_glue('Number of rows in data imported: {nrow(lvdreport)}'), '\n'))

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
rawepichr2 <- rbind(c("Heart Removed Due to Death","","","",""), rawepichr)
kk <- addWorksheet(wb, sheetName = "Kidney")
rawepickk2 <- rbind(c("Kidney Removed Due to Death","","","",""), rawepickk)
li <- addWorksheet(wb, sheetName = "Liver")
rawepicli2 <- rbind(c("Liver Removed Due to Death","","","",""), rawepicli)
panc <- addWorksheet(wb, sheetName = "Pancreas")
rawepicpanc2 <- rbind(c("Pancreas Removed Due to Death","","","",""), rawepicpanc)

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
colnames(rawepic) = gsub("Translant.Evaluation.Start.DS", "Eval.Start", colnames(rawepic))
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
rawepic2 <- rawepic %>%
	select(-c("Visit.Type.NM","Transplant.Episode.Record.ID","Min..Translant.Evaluation.Start.DS","X","Appointment.Status.NM"))
write_xlsx(rawepic2, "LVD Report Print.xlsx")
month18filelist<- list.files(pattern = "LVD Report Print", full.names = TRUE, ignore.case = TRUE)
rawepic <- as.data.frame(read_excel(month18filelist))
rawepic <- rbind(c("LVD Report","","","","","",""), rawepic)
wb <- loadWorkbook(month18filelist)

sheet_names <- names(wb)
for (sheet in sheet_names) {
print(sheet)
freezePane(wb,sheet =sheet, firstRow = TRUE, firstCol =FALSE)
addFilter(wb, sheet =sheet, rows = 1, cols =1:ncol(read.xlsx(wb,sheet)))
setColWidths(wb, sheet =sheet, cols =1:ncol(read.xlsx(wb,sheet)), widths= c(38,11,11,12,11,14,11))
}
saveWorkbook(wb, "LVD Report Print.xlsx", overwrite =TRUE)
# move raw files away
dir.create(file.path(todaydirraw))
dir.create(file.path(todaydirprint))
rawfiles1 <- list.files(pattern = ".*raw.*\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
rawfiles2 <- list.files(pattern = "TXP_CMS", full.names = TRUE, ignore.case = TRUE)
darfiles <- list.files(pattern = "DAR", full.names = TRUE, ignore.case = TRUE)

for (file in darfiles) {
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

#making pdfs
readline('right click on the creatCMSpdfs script, open in powershell \n script will ask for two directories you can paste from address bar\nuse the same dir twice if needed\n')
clipr::write_clip(todaydir)
shell.exec(r'(M:\PAH\Transplant Analyst Data\Analyst Documentation\Automated\DailyCMSReport)')
readline('path of source dir in clipboard, just paste in powershell\npress enter in R to load target dir to clipboard \n')
clipr::write_clip(todaydirprint)
readline('target dir path in clipboard, it will take a bit to run,\npowershell will close when finished. This is the last step and will take a min or two...\n')
cat('CMS daily reports completed, great job\n')



