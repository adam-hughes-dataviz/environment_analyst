list.of.packages <- c("xlsx","dplyr","tidyverse", "readxl")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

wd <- setwd("C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data")


      ###   AEC   ###

  AEC_data <- read.xlsx("G-AEC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(AEC_data) = make.names(AEC_data$AEC, unique = T)
  AEC_data <- tibble::rownames_to_column(AEC_data,"Variable")
  AEC_data <- select(AEC_data, -2)
  
  AEC_data_reshaped <- reshape(AEC_data,
                               direction = "long",
                               varying = list(names(AEC_data)[2:17]),
                               v.names = "Value",
                               idvar = "AEC",
                               timevar = "Year",
                               times = 2004:2019)
  
  AEC_data_reshaped <- select(AEC_data_reshaped, -4)
  AEC_data_reshaped$Company <- "AEC"
  
  write.xlsx(AEC_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-AEC-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  
      ###   ANT   ###
  
  ANT_data <- read.xlsx("G-ANT-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(ANT_data) = make.names(ANT_data$ANT, unique = T)
  ANT_data <- tibble::rownames_to_column(ANT_data,"Variable")
  ANT_data <- select(ANT_data, -2)
  
  ANT_data_reshaped <- reshape(ANT_data,
                               direction = "long",
                               varying = list(names(ANT_data)[2:17]),
                               v.names = "Value",
                               idvar = "ANT",
                               timevar = "Year",
                               times = 2004:2019)
  
  ANT_data_reshaped <- select(ANT_data_reshaped, -4)
  ANT_data_reshaped$Company <- "ANT"
  
  write.xlsx(ANT_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-ANT-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  
      ###   ARC   ###
  
  ARC_data <- read.xlsx("G-ARC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(ARC_data) = make.names(ARC_data$ARC, unique = T)
  ARC_data <- tibble::rownames_to_column(ARC_data,"Variable")
  ARC_data <- select(ARC_data, -2)
  
  ARC_data_reshaped <- reshape(ARC_data,
                               direction = "long",
                               varying = list(names(ARC_data)[2:17]),
                               v.names = "Value",
                               idvar = "ARC",
                               timevar = "Year",
                               times = 2004:2019)
  
  ARC_data_reshaped <- select(ARC_data_reshaped, -4)
  ARC_data_reshaped$Company <- "ARC"
  
  write.xlsx(ARC_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-ARC-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   CAR   ###
  
  CAR_data <- read.xlsx("G-CAR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(CAR_data) = make.names(CAR_data$CAR, unique = T)
  CAR_data <- tibble::rownames_to_column(CAR_data,"Variable")
  CAR_data <- select(CAR_data, -2)
  
  CAR_data_reshaped <- reshape(CAR_data,
                               direction = "long",
                               varying = list(names(CAR_data)[2:17]),
                               v.names = "Value",
                               idvar = "CAR",
                               timevar = "Year",
                               times = 2004:2019)
  
  CAR_data_reshaped <- select(CAR_data_reshaped, -4)
  CAR_data_reshaped$Company <- "CAR"
  
  write.xlsx(CAR_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-CAR-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   ERM   ###
  
  ERM_data <- read.xlsx("G-ERM-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(ERM_data) = make.names(ERM_data$ERM, unique = T)
  ERM_data <- tibble::rownames_to_column(ERM_data,"Variable")
  ERM_data <- select(ERM_data, -2)
  
  ERM_data_reshaped <- reshape(ERM_data,
                               direction = "long",
                               varying = list(names(ERM_data)[2:17]),
                               v.names = "Value",
                               idvar = "ERM",
                               timevar = "Year",
                               times = 2004:2019)
  
  ERM_data_reshaped <- select(ERM_data_reshaped, -4)
  ERM_data_reshaped$Company <- "ERM"
  
  write.xlsx(ERM_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-ERM-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   GHD   ###
  
  GHD_data <- read.xlsx("G-GHD-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(GHD_data) = make.names(GHD_data$GHD, unique = T)
  GHD_data <- tibble::rownames_to_column(GHD_data,"Variable")
  GHD_data <- select(GHD_data, -2)
  
  GHD_data_reshaped <- reshape(GHD_data,
                               direction = "long",
                               varying = list(names(GHD_data)[2:17]),
                               v.names = "Value",
                               idvar = "GHD",
                               timevar = "Year",
                               times = 2004:2019)
  
  GHD_data_reshaped <- select(GHD_data_reshaped, -4)
  GHD_data_reshaped$Company <- "GHD"
  
  write.xlsx(GHD_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-GHD-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   GOL   ###
  
  GOL_data <- read.xlsx("G-GOL-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(GOL_data) = make.names(GOL_data$GOL, unique = T)
  GOL_data <- tibble::rownames_to_column(GOL_data,"Variable")
  GOL_data <- select(GOL_data, -2)
  
  GOL_data_reshaped <- reshape(GOL_data,
                               direction = "long",
                               varying = list(names(GOL_data)[2:17]),
                               v.names = "Value",
                               idvar = "GOL",
                               timevar = "Year",
                               times = 2004:2019)
  
  GOL_data_reshaped <- select(GOL_data_reshaped, -4)
  GOL_data_reshaped$Company <- "GOL"
  
  write.xlsx(GOL_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-GOL-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   HDR   ###
  
  HDR_data <- read.xlsx("G-HDR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(HDR_data) = make.names(HDR_data$HDR, unique = T)
  HDR_data <- tibble::rownames_to_column(HDR_data,"Variable")
  HDR_data <- select(HDR_data, -2)
  
  HDR_data_reshaped <- reshape(HDR_data,
                               direction = "long",
                               varying = list(names(HDR_data)[2:17]),
                               v.names = "Value",
                               idvar = "HDR",
                               timevar = "Year",
                               times = 2004:2019)
  
  HDR_data_reshaped <- select(HDR_data_reshaped, -4)
  HDR_data_reshaped$Company <- "HDR"
  
  write.xlsx(HDR_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-HDR-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   ICF   ###
  
  ICF_data <- read.xlsx("G-ICF-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(ICF_data) = make.names(ICF_data$ICF, unique = T)
  ICF_data <- tibble::rownames_to_column(ICF_data,"Variable")
  ICF_data <- select(ICF_data, -2)
  
  ICF_data_reshaped <- reshape(ICF_data,
                               direction = "long",
                               varying = list(names(ICF_data)[2:17]),
                               v.names = "Value",
                               idvar = "ICF",
                               timevar = "Year",
                               times = 2004:2019)
  
  ICF_data_reshaped <- select(ICF_data_reshaped, -4)
  ICF_data_reshaped$Company <- "ICF"
  
  write.xlsx(ICF_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-ICF-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   INO   ###
  
  INO_data <- read.xlsx("G-INO-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(INO_data) = make.names(INO_data$INO, unique = T)
  INO_data <- tibble::rownames_to_column(INO_data,"Variable")
  INO_data <- select(INO_data, -2)
  
  INO_data_reshaped <- reshape(INO_data,
                               direction = "long",
                               varying = list(names(INO_data)[2:17]),
                               v.names = "Value",
                               idvar = "INO",
                               timevar = "Year",
                               times = 2004:2019)
  
  INO_data_reshaped <- select(INO_data_reshaped, -4)
  INO_data_reshaped$Company <- "INO"
  
  write.xlsx(INO_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-INO-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   JAC   ###
  
  JAC_data <- read.xlsx("G-JAC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(JAC_data) = make.names(JAC_data$JAC, unique = T)
  JAC_data <- tibble::rownames_to_column(JAC_data,"Variable")
  JAC_data <- select(JAC_data, -2)
  
  JAC_data_reshaped <- reshape(JAC_data,
                               direction = "long",
                               varying = list(names(JAC_data)[2:17]),
                               v.names = "Value",
                               idvar = "JAC",
                               timevar = "Year",
                               times = 2004:2019)
  
  JAC_data_reshaped <- select(JAC_data_reshaped, -4)
  JAC_data_reshaped$Company <- "JAC"
  
  write.xlsx(JAC_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-JAC-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   MOT   ###
  
  MOT_data <- read.xlsx("G-MOT-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(MOT_data) = make.names(MOT_data$MOT, unique = T)
  MOT_data <- tibble::rownames_to_column(MOT_data,"Variable")
  MOT_data <- select(MOT_data, -2)
  
  MOT_data_reshaped <- reshape(MOT_data,
                               direction = "long",
                               varying = list(names(MOT_data)[2:17]),
                               v.names = "Value",
                               idvar = "MOT",
                               timevar = "Year",
                               times = 2004:2019)
  
  MOT_data_reshaped <- select(MOT_data_reshaped, -4)
  MOT_data_reshaped$Company <- "MOT"
  
  write.xlsx(MOT_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-MOT-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   RAM   ###
  
  RAM_data <- read.xlsx("G-RAM-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(RAM_data) = make.names(RAM_data$RAM, unique = T)
  RAM_data <- tibble::rownames_to_column(RAM_data,"Variable")
  RAM_data <- select(RAM_data, -2)
  
  RAM_data_reshaped <- reshape(RAM_data,
                               direction = "long",
                               varying = list(names(RAM_data)[2:17]),
                               v.names = "Value",
                               idvar = "RAM",
                               timevar = "Year",
                               times = 2004:2019)
  
  RAM_data_reshaped <- select(RAM_data_reshaped, -4)
  RAM_data_reshaped$Company <- "RAM"
  
  write.xlsx(RAM_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-RAM-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   RHD   ###
  
  RHD_data <- read.xlsx("G-RHD-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(RHD_data) = make.names(RHD_data$RHD, unique = T)
  RHD_data <- tibble::rownames_to_column(RHD_data,"Variable")
  RHD_data <- select(RHD_data, -2)
  
  RHD_data_reshaped <- reshape(RHD_data,
                               direction = "long",
                               varying = list(names(RHD_data)[2:17]),
                               v.names = "Value",
                               idvar = "RHD",
                               timevar = "Year",
                               times = 2004:2019)
  
  RHD_data_reshaped <- select(RHD_data_reshaped, -4)
  RHD_data_reshaped$Company <- "RHD"
  
  write.xlsx(RHD_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-RHD-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   RPS   ###
  
  RPS_data <- read.xlsx("G-RPS-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(RPS_data) = make.names(RPS_data$RPS, unique = T)
  RPS_data <- tibble::rownames_to_column(RPS_data,"Variable")
  RPS_data <- select(RPS_data, -2)
  
  RPS_data_reshaped <- reshape(RPS_data,
                               direction = "long",
                               varying = list(names(RPS_data)[2:17]),
                               v.names = "Value",
                               idvar = "RPS",
                               timevar = "Year",
                               times = 2004:2019)
  
  RPS_data_reshaped <- select(RPS_data_reshaped, -4)
  RPS_data_reshaped$Company <- "RPS"
  
  write.xlsx(RPS_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-RPS-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   SLR   ###
  
  SLR_data <- read.xlsx("G-SLR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(SLR_data) = make.names(SLR_data$SLR, unique = T)
  SLR_data <- tibble::rownames_to_column(SLR_data,"Variable")
  SLR_data <- select(SLR_data, -2)
  
  SLR_data_reshaped <- reshape(SLR_data,
                               direction = "long",
                               varying = list(names(SLR_data)[2:17]),
                               v.names = "Value",
                               idvar = "SLR",
                               timevar = "Year",
                               times = 2004:2019)
  
  SLR_data_reshaped <- select(SLR_data_reshaped, -4)
  SLR_data_reshaped$Company <- "SLR"
  
  write.xlsx(SLR_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-SLR-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   SME   ###
  
  SME_data <- read.xlsx("G-SME-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(SME_data) = make.names(SME_data$SME, unique = T)
  SME_data <- tibble::rownames_to_column(SME_data,"Variable")
  SME_data <- select(SME_data, -2)
  
  SME_data_reshaped <- reshape(SME_data,
                               direction = "long",
                               varying = list(names(SME_data)[2:17]),
                               v.names = "Value",
                               idvar = "SME",
                               timevar = "Year",
                               times = 2004:2019)
  
  SME_data_reshaped <- select(SME_data_reshaped, -4)
  SME_data_reshaped$Company <- "SME"
  
  write.xlsx(SME_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-SME-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   SNC   ###
  
  SNC_data <- read.xlsx("G-SNC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(SNC_data) = make.names(SNC_data$SNC, unique = T)
  SNC_data <- tibble::rownames_to_column(SNC_data,"Variable")
  SNC_data <- select(SNC_data, -2)
  
  SNC_data_reshaped <- reshape(SNC_data,
                               direction = "long",
                               varying = list(names(SNC_data)[2:17]),
                               v.names = "Value",
                               idvar = "SNC",
                               timevar = "Year",
                               times = 2004:2019)
  
  SNC_data_reshaped <- select(SNC_data_reshaped, -4)
  SNC_data_reshaped$Company <- "SNC"
  
  write.xlsx(SNC_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-SNC-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   STA   ###
  
  STA_data <- read.xlsx("G-STA-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(STA_data) = make.names(STA_data$STA, unique = T)
  STA_data <- tibble::rownames_to_column(STA_data,"Variable")
  STA_data <- select(STA_data, -2)
  
  STA_data_reshaped <- reshape(STA_data,
                               direction = "long",
                               varying = list(names(STA_data)[2:17]),
                               v.names = "Value",
                               idvar = "STA",
                               timevar = "Year",
                               times = 2004:2019)
  
  STA_data_reshaped <- select(STA_data_reshaped, -4)
  STA_data_reshaped$Company <- "STA"
  
  write.xlsx(STA_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-STA-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   SWE   ###
  
  SWE_data <- read.xlsx("G-SWE-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(SWE_data) = make.names(SWE_data$SWE, unique = T)
  SWE_data <- tibble::rownames_to_column(SWE_data,"Variable")
  SWE_data <- select(SWE_data, -2)
  
  SWE_data_reshaped <- reshape(SWE_data,
                               direction = "long",
                               varying = list(names(SWE_data)[2:17]),
                               v.names = "Value",
                               idvar = "SWE",
                               timevar = "Year",
                               times = 2004:2019)
  
  SWE_data_reshaped <- select(SWE_data_reshaped, -4)
  SWE_data_reshaped$Company <- "SWE"
  
  write.xlsx(SWE_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-SWE-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   TAU   ###
  
  TAU_data <- read.xlsx("G-TAU-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(TAU_data) = make.names(TAU_data$TAU, unique = T)
  TAU_data <- tibble::rownames_to_column(TAU_data,"Variable")
  TAU_data <- select(TAU_data, -2)
  
  TAU_data_reshaped <- reshape(TAU_data,
                               direction = "long",
                               varying = list(names(TAU_data)[2:17]),
                               v.names = "Value",
                               idvar = "TAU",
                               timevar = "Year",
                               times = 2004:2019)
  
  TAU_data_reshaped <- select(TAU_data_reshaped, -4)
  TAU_data_reshaped$Company <- "TAU"
  
  write.xlsx(TAU_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-TAU-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   TET   ###
  
  TET_data <- read.xlsx("G-TET-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(TET_data) = make.names(TET_data$TET, unique = T)
  TET_data <- tibble::rownames_to_column(TET_data,"Variable")
  TET_data <- select(TET_data, -2)
  
  TET_data_reshaped <- reshape(TET_data,
                               direction = "long",
                               varying = list(names(TET_data)[2:17]),
                               v.names = "Value",
                               idvar = "TET",
                               timevar = "Year",
                               times = 2004:2019)
  
  TET_data_reshaped <- select(TET_data_reshaped, -4)
  TET_data_reshaped$Company <- "TET"
  
  write.xlsx(TET_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-TET-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   WOO   ###
  
  WOO_data <- read.xlsx("G-WOO-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(WOO_data) = make.names(WOO_data$WOO, unique = T)
  WOO_data <- tibble::rownames_to_column(WOO_data,"Variable")
  WOO_data <- select(WOO_data, -2)
  
  WOO_data_reshaped <- reshape(WOO_data,
                               direction = "long",
                               varying = list(names(WOO_data)[2:17]),
                               v.names = "Value",
                               idvar = "WOO",
                               timevar = "Year",
                               times = 2004:2019)
  
  WOO_data_reshaped <- select(WOO_data_reshaped, -4)
  WOO_data_reshaped$Company <- "WOO"
  
  write.xlsx(WOO_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-WOO-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  ###   WOR   ###
  
  WOR_data <- read.xlsx("G-WOR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(WOR_data) = make.names(WOR_data$WOR, unique = T)
  WOR_data <- tibble::rownames_to_column(WOR_data,"Variable")
  WOR_data <- select(WOR_data, -2)
  
  WOR_data_reshaped <- reshape(WOR_data,
                               direction = "long",
                               varying = list(names(WOR_data)[2:17]),
                               v.names = "Value",
                               idvar = "WOR",
                               timevar = "Year",
                               times = 2004:2019)
  
  WOR_data_reshaped <- select(WOR_data_reshaped, -4)
  WOR_data_reshaped$Company <- "WOR"
  
  write.xlsx(WOR_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-WOR-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  
  ###   WSP   ###
  
  WSP_data <- read.xlsx("G-WSP-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
  
  rownames(WSP_data) = make.names(WSP_data$WSP, unique = T)
  WSP_data <- tibble::rownames_to_column(WSP_data,"Variable")
  WSP_data <- select(WSP_data, -2)
  
  WSP_data_reshaped <- reshape(WSP_data,
                               direction = "long",
                               varying = list(names(WSP_data)[2:17]),
                               v.names = "Value",
                               idvar = "WSP",
                               timevar = "Year",
                               times = 2004:2019)
  
  WSP_data_reshaped <- select(WSP_data_reshaped, -4)
  WSP_data_reshaped$Company <- "WSP"
  
  write.xlsx(WSP_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-WSP-DATA-LONG.xlsx", row.names = F, showNA = F)
  
  
  
  ### BIND ###
  
combined_data <- rbind(AEC_data_reshaped,ANT_data_reshaped,ARC_data_reshaped,CAR_data_reshaped,ERM_data_reshaped,GHD_data_reshaped,GOL_data_reshaped,
                       HDR_data_reshaped,ICF_data_reshaped,INO_data_reshaped,JAC_data_reshaped,MOT_data_reshaped,RAM_data_reshaped,RHD_data_reshaped,
                       RPS_data_reshaped,SLR_data_reshaped,SME_data_reshaped,SNC_data_reshaped,STA_data_reshaped,SWE_data_reshaped,TAU_data_reshaped,
                       TET_data_reshaped,WOO_data_reshaped,WOR_data_reshaped,WSP_data_reshaped)
write.xlsx(combined_data, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/Z-Rbind-UK-Profiles.xlsx", row.names = F, showNA = F)

  
      
