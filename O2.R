library(dplyr)
library(tidyr)
library(data.table)
library(readxl)
library(lubridate)
library(do)
library(VIM)
library(AER) #邏輯斯用
library(IDPmisc)#去掉Inf用

#setwd('D:/user/研究所/金融市場')
options(scipen=999)
#處理CPI和破產資料-----------------------------------
#CPI
idx_Fin <- strftime(seq.Date(as.Date('1995-12-01'), as.Date('2020-12-01'), by='month'), format = '%Y%m')
CPI <- read_excel("D:/user/研究所/蹤影/論文/TEJ download(~202012)/CPI_202012.xlsx")
colnames(CPI) <- c("Date","CPI","Predict")
CPI <- CPI[-(1:2),-3]
CPI$Date <- openxlsx::convertToDate(CPI$Date)
CPI$Date <- paste0(substr(CPI$Date,1,4),substr(CPI$Date,6,7)) %>% as.numeric()
CPI$CPI <- CPI$CPI %>% as.numeric()
CPI <- CPI %>% setDT
CPI <- CPI[substr(Date,5,6) == '12']

#上下市破產資料
Companybankrupt <- fread("D:/user/研究所/金融市場/bankrupt.txt") %>% 
  rename(.,Company="證券代碼",State = '目前狀態',Company2 = '公司中文簡稱',IPO = "公開發行日",
         TSE="上市日期",UNTSE="下市日期",default="危機事件類別",Caption = '危機事件類別說明') %>% 
  separate(.,Caption,c('Caption','defDate'),'\\(') %>% 
  mutate(Company = gsub(' ','', Company),default = ifelse(default == 'D'|default == 'H'|default == 'C'|default == 'G'|default == 'I'|default == 'E'|
                                                            default == 'Z'|default == 'N'|default == 'S',1,0),defDate = substr(defDate,1,4)) %>% 
  .[,.(Company,default)]#,Date = ifelse(Date =='非5',NA,Date) %>% as.numeric())] %>% na.omit()

#處理財報------------------------------------------------------
#上市櫃
TSE <- fread("D:/user/研究所/金融市場/tej/TSE.txt", colClasses = 'numeric') %>% 
  rename(.,Company="證券代碼",Type="TSE 產業別",Date="年月",CA="流動資產",CL="流動負債",TL="負債總額",TE="股東權益總額",
         TA="資產總額",NI="繼續營業單位損益",TP="當期所得稅負債") %>% 
  mutate(Company = gsub(' ','', Company), Date = Date %>% as.numeric) %>% 
  arrange(Company,desc(Date)) %>% .[!Type %in% c(14,17) & substr(Date,5,6) == '12']

TSE$default <- 0

#下市櫃
UNTSE <- fread("D:/user/研究所/金融市場/tej/UNTSE.txt", colClasses = 'numeric') %>% 
  rename(.,Company="證券代碼",Type="TSE 產業別",Date="年月",CA="流動資產",CL="流動負債",TL="負債總額",TE="股東權益總額",
         TA="資產總額",NI="繼續營業單位損益",TP="當期所得稅負債") %>% 
  mutate(Company = gsub(' ','', Company), Date = Date %>% as.numeric) %>% 
  arrange(Company,desc(Date)) %>% .[!Type %in% c(14,17) & substr(Date,5,6) == '12']
  
dt_1 <- UNTSE %>% .[,.SD[1],by = Company]
UNTSE_2 <- left_join(dt_1,Companybankrupt) %>% mutate_at(.,'default', ~replace(., is.na(.), 0)) %>% 
  .[,.(Company,Date,default)]

UNTSE_fin <- left_join(UNTSE,UNTSE_2) %>% mutate_at(.,'default', ~replace(., is.na(.), 0))

All_acc <- rbind(TSE,UNTSE_fin) %>% group_by(Company) %>% 
  fill(TL:TP,.direction = 'down') %>% setDT %>% 
  .[,NI2 := lead(NI,1),by = Company] %>%  .[-which(is.na(.[['CA']]) & is.na(.[['CL']]))]

#樣本內
Insample <- All_acc %>% .[Date %between% c(200012,201012),.(Company,Date,CA,CL,TL,TE,TA,NI,TP,NI2,default)] %>% na.omit
Insample <- left_join(Insample,CPI)

#變數
all_variable <- Insample %>% .[,.(Date = Date, B1 = log(TA/1000)/CPI, B2 = TL/TA,
                                      B3 = (CA-CL)/TA, B4 = CL/CA, B5 = ifelse(TL > TA, 1, 0),
                                      B6 = NI/TA, B7 = (TP+NI)/TL, B8 = ifelse((NI < 0 & NI2 < 0), 1, 0),
                                      B9 = (NI-NI2)/(abs(NI)+abs(NI2)),default = default), by = Company] %>% NaRV.omit()

myfit <- glm(default ~ B1 + B2 + B3 + B4 + B5 + B6 + B7 + B8 + B9, data=all_variable, family=binomial())
summary(myfit)


#樣本外預估
OutSample <- All_acc %>% .[Date %between% c(201012,202012),.(Company,Date,CA,CL,TL,TE,TA,NI,TP,NI2,default)] %>% na.omit
OutSample <- left_join(OutSample,CPI) %>% 
  .[,.(Date = Date, B1 = log(TA/1000)/CPI, B2 = TL/TA,
       B3 = (CA-CL)/TA, B4 = CL/CA, B5 = ifelse(TL > TA, 1, 0),
       B6 = NI/TA, B7 = (TP+NI)/TL, B8 = ifelse((NI < 0 & NI2 < 0), 1, 0),
       B9 = (NI-NI2)/(abs(NI)+abs(NI2)),default), by = Company] %>% 
  .[,Oscore := -9.1939595
    + 13.9504872 * B1
    + 0.0359081 * B2
    + 0.0367629 * B3
    - 0.0005516 * B4
    + 1.9597434 * B5
    - 0.5487264 * B6
    + 0.0941578 * B7
    + 3.3662881 * B8
    - 0.1485644 * B9] %>% 
  .[,P_default := exp(Oscore)/(1+exp(Oscore))] %>% 
  .[,lag_def := lead(P_default,1),by = Company] %>% na.omit

predict <- OutSample %>% 
  mutate(Predict_50 = ifelse(.[['default']] == 1 & .[['lag_def']] >= 0.5,1,ifelse(.[['default']] == 0 & .[['lag_def']] <= 0.5,1,0)),
         Predict_70 = ifelse(.[['default']] == 1 & .[['lag_def']] >= 0.7,1,ifelse(.[['default']] == 0 & .[['lag_def']] <= 0.7,1,0)),
         Predict_90 = ifelse(.[['default']] == 1 & .[['lag_def']] >= 0.9,1,ifelse(.[['default']] == 0 & .[['lag_def']] <= 0.9,1,0)))

df_oscore <- data.frame(Oscore_50 = c(nrow(predict),sum(predict[['Predict_50']]),sum(predict[['Predict_50']])/nrow(predict)),
                        Oscore_70 = c(nrow(predict),sum(predict[['Predict_70']]),sum(predict[['Predict_70']])/nrow(predict)),
                        Oscore_90 = c(nrow(predict),sum(predict[['Predict_90']]),sum(predict[['Predict_90']])/nrow(predict)))

row.names(df_oscore) <- c('總樣本數','準確數','準確率')

write.csv(df_oscore,"Oscore_predict_TW.CSV", row.names = TRUE)

aggr(OutSample,prop = FALSE,number = TRUE)










