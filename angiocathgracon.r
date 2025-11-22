#angiograms for gracon
#from "M:\PAH\Transplant Analyst Data\duyXQ\angiogram post txp.twbx"
# paste txp detail to 5ybase adds new ones
sheetpath1 <- r'(M:\PAH\Transplant Analyst Data\Kidney misc jada\Jada\Adhoc\Gracon\txps with cardiac caths- angios last 5Ybase.xlsx)'
angio2 <- importXLsheets(sheetpath1)
angio <- angio2$txpdetailrawpaste %>% filter(`Tx Episode Organ UNOS Definition` == "Kidney")
angio$MRN <- as.character(angio$MRN)
str(angio)

sheetpath2 <- r'(M:\PAH\Transplant Analyst Data\Kidney misc jada\Jada\Adhoc\Gracon\sCrallorgans__remaining_at_risk_20251104_2058.xlsx)'
scr2 <- importXLsheets(sheetpath2)
scr <- scr2$Sheet1
scr$MRN <- as.character(scr$MRN)
scr<- scr %>% select(c("MRN","DOB","Last CR","Last CR Date"))
str(scr)

final_data <- angio %>%
  left_join(scr, by = "MRN") %>%
  mutate(TxptoAngio = `Day of Angiogram Procedure Start DTS` - `Transplant Surgery DTS`) %>%
  select(c("Transplant Surgery DTS","MRN","Day of Patient Birth DTS","Tx Episode Organ UNOS Definition","Primary Log Procedure NM","Last CR","Last CR Date","Transplant Discharge DTS","Day of Angiogram Procedure Start DTS","TxptoAngio","All Case Procedures TXT"))

clipr::write_clip(final_data)
todaydir<- Sys.Date()
write_xlsx(final_data, str_glue(r'(M:\PAH\Transplant Analyst Data\Kidney misc jada\Jada\Adhoc\Gracon\txpcardiaccaths{todaydir}.xlsx)'))