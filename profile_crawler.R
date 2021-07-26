#Install packages - Ensure No. TRUE = No. packages
list.of.packages <- c("xlsx","tidyverse","tibble")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

#set working directory
setwd("C:/Users/adamh/Google Drive/Global MIS/Global profiles")

###   1. AEC   ###

#read in company profile data tab and clean to filter rows and columns
AEC_data_global <- read.xlsx("G-AEC.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
AEC_col_clean <- AEC_data_global[,-2]
AEC_row_clean <- AEC_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
AEC_G_no_nas <- AEC_row_clean[rowSums(is.na(AEC_row_clean)) !=ncol(AEC_row_clean), ]
#Change cell A1 to the company TLA
names(AEC_G_no_nas)[1] <-"AEC"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(AEC_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-AEC-DATA.xlsx", row.names = F, showNA = F)

###   2. ANT   ###

#read in company profile data tab and clean to filter rows and columns
ANT_data_global <- read.xlsx("G-ANT.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
ANT_col_clean <- ANT_data_global[,-2]
ANT_row_clean <- ANT_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
ANT_G_no_nas <- ANT_row_clean[rowSums(is.na(ANT_row_clean)) !=ncol(ANT_row_clean), ]
#Change cell A1 to the company TLA
names(ANT_G_no_nas)[1] <-"ANT"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(ANT_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-ANT-DATA.xlsx", row.names = F, showNA = F)

###  3. ARC   ###

#read in company profile data tab and clean to filter rows and columns
ARC_data_global <- read.xlsx("G-ARC.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
ARC_col_clean <- ARC_data_global[,-2]
ARC_row_clean <- ARC_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
ARC_G_no_nas <- ARC_row_clean[rowSums(is.na(ARC_row_clean)) !=ncol(ARC_row_clean), ]
#Change cell A1 to the company TLA
names(ARC_G_no_nas)[1] <-"ARC"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(ARC_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-ARC-DATA.xlsx", row.names = F, showNA = F)

###   4. CAR   ###

#read in company profile data tab and clean to filter rows and columns
CAR_data_global <- read.xlsx("G-CAR.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
CAR_col_clean <- CAR_data_global[,-2]
CAR_row_clean <- CAR_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
CAR_G_no_nas <- CAR_row_clean[rowSums(is.na(CAR_row_clean)) !=ncol(CAR_row_clean), ]
#Change cell A1 to the company TLA
names(CAR_G_no_nas)[1] <-"CAR"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(CAR_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-CAR-DATA.xlsx", row.names = F, showNA = F)

###   5. ERM   ###

#read in company profile data tab and clean to filter rows and columns
ERM_data_global <- read.xlsx("G-ERM.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
ERM_col_clean <- ERM_data_global[,-2]
ERM_row_clean <- ERM_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
ERM_G_no_nas <- ERM_row_clean[rowSums(is.na(ERM_row_clean)) !=ncol(ERM_row_clean), ]
#Change cell A1 to the company TLA
names(ERM_G_no_nas)[1] <-"ERM"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(ERM_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-ERM-DATA.xlsx", row.names = F, showNA = F)

###   6. GHD   ###

#read in company profile data tab and clean to filter rows and columns
GHD_data_global <- read.xlsx("G-GHD.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
GHD_col_clean <- GHD_data_global[,-2]
GHD_row_clean <- GHD_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
GHD_G_no_nas <- GHD_row_clean[rowSums(is.na(GHD_row_clean)) !=ncol(GHD_row_clean), ]
#Change cell A1 to the company TLA
names(GHD_G_no_nas)[1] <-"GHD"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(GHD_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-GHD-DATA.xlsx", row.names = F, showNA = F)

###   7. GOL   ###

#read in company profile data tab and clean to filter rows and columns
GOL_data_global <- read.xlsx("G-GOL.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
GOL_col_clean <- GOL_data_global[,-2]
GOL_row_clean <- GOL_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
GOL_G_no_nas <- GOL_row_clean[rowSums(is.na(GOL_row_clean)) !=ncol(GOL_row_clean), ]
#Change cell A1 to the company TLA
names(GOL_G_no_nas)[1] <-"GOL"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(GOL_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-GOL-DATA.xlsx", row.names = F, showNA = F)

###   8. HDR   ###

#read in company profile data tab and clean to filter rows and columns
HDR_data_global <- read.xlsx("G-HDR.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
HDR_col_clean <- HDR_data_global[,-2]
HDR_row_clean <- HDR_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
HDR_G_no_nas <- HDR_row_clean[rowSums(is.na(HDR_row_clean)) !=ncol(HDR_row_clean), ]
#Change cell A1 to the company TLA
names(HDR_G_no_nas)[1] <-"HDR"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(HDR_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-HDR-DATA.xlsx", row.names = F, showNA = F)

###  9. ICF   ###

#read in company profile data tab and clean to filter rows and columns
ICF_data_global <- read.xlsx("G-ICF.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
ICF_col_clean <- ICF_data_global[,-2]
ICF_row_clean <- ICF_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
ICF_G_no_nas <- ICF_row_clean[rowSums(is.na(ICF_row_clean)) !=ncol(ICF_row_clean), ]
#Change cell A1 to the company TLA
names(ICF_G_no_nas)[1] <-"ICF"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(ICF_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-ICF-DATA.xlsx", row.names = F, showNA = F)

###   10. INO   ###

#read in company profile data tab and clean to filter rows and columns
INO_data_global <- read.xlsx("G-INO.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
INO_col_clean <- INO_data_global[,-2]
INO_row_clean <- INO_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
INO_G_no_nas <- INO_row_clean[rowSums(is.na(INO_row_clean)) !=ncol(INO_row_clean), ]
#Change cell A1 to the company TLA
names(INO_G_no_nas)[1] <-"INO"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(INO_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-INO-DATA.xlsx", row.names = F, showNA = F)

###   11. JAC   ###

#read in company profile data tab and clean to filter rows and columns
JAC_data_global <- read.xlsx("G-JAC.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
JAC_col_clean <- JAC_data_global[,-2]
JAC_row_clean <- JAC_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
JAC_G_no_nas <- JAC_row_clean[rowSums(is.na(JAC_row_clean)) !=ncol(JAC_row_clean), ]
#Change cell A1 to the company TLA
names(JAC_G_no_nas)[1] <-"JAC"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(JAC_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-JAC-DATA.xlsx", row.names = F, showNA = F)

###   12. MOT   ###

#read in company profile data tab and clean to filter rows and columns
MOT_data_global <- read.xlsx("G-MOT.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
MOT_col_clean <- MOT_data_global[,-2]
MOT_row_clean <- MOT_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
MOT_G_no_nas <- MOT_row_clean[rowSums(is.na(MOT_row_clean)) !=ncol(MOT_row_clean), ]
#Change cell A1 to the company TLA
names(MOT_G_no_nas)[1] <-"MOT"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(MOT_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-MOT-DATA.xlsx", row.names = F, showNA = F)

###   13. RAM   ###

#read in company profile data tab and clean to filter rows and columns
RAM_data_global <- read.xlsx("G-RAM.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
RAM_col_clean <- RAM_data_global[,-2]
RAM_row_clean <- RAM_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
RAM_G_no_nas <- RAM_row_clean[rowSums(is.na(RAM_row_clean)) !=ncol(RAM_row_clean), ]
#Change cell A1 to the company TLA
names(RAM_G_no_nas)[1] <-"RAM"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(RAM_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-RAM-DATA.xlsx", row.names = F, showNA = F)

###   14. RHD   ###

#read in company profile data tab and clean to filter rows and columns
RHD_data_global <- read.xlsx("G-RHD.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
RHD_col_clean <- RHD_data_global[,-2]
RHD_row_clean <- RHD_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
RHD_G_no_nas <- RHD_row_clean[rowSums(is.na(RHD_row_clean)) !=ncol(RHD_row_clean), ]
#Change cell A1 to the company TLA
names(RHD_G_no_nas)[1] <-"RHD"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(RHD_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-RHD-DATA.xlsx", row.names = F, showNA = F)

###   15. RPS   ###

#read in company profile data tab and clean to filter rows and columns
RPS_data_global <- read.xlsx("G-RPS.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
RPS_col_clean <- RPS_data_global[,-2]
RPS_row_clean <- RPS_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
RPS_G_no_nas <- RPS_row_clean[rowSums(is.na(RPS_row_clean)) !=ncol(RPS_row_clean), ]
#Change cell A1 to the company TLA
names(RPS_G_no_nas)[1] <-"RPS"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(RPS_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-RPS-DATA.xlsx", row.names = F, showNA = F)

###  16. SLR   ###

#read in company profile data tab and clean to filter rows and columns
SLR_data_global <- read.xlsx("G-SLR.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
SLR_col_clean <- SLR_data_global[,-2]
SLR_row_clean <- SLR_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
SLR_G_no_nas <- SLR_row_clean[rowSums(is.na(SLR_row_clean)) !=ncol(SLR_row_clean), ]
#Change cell A1 to the company TLA
names(SLR_G_no_nas)[1] <-"SLR"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(SLR_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-SLR-DATA.xlsx", row.names = F, showNA = F)

###  17. SME   ###

#read in company profile data tab and clean to filter rows and columns
SME_data_global <- read.xlsx("G-SME.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
SME_col_clean <- SME_data_global[,-2]
SME_row_clean <- SME_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
SME_G_no_nas <- SME_row_clean[rowSums(is.na(SME_row_clean)) !=ncol(SME_row_clean), ]
#Change cell A1 to the company TLA
names(SME_G_no_nas)[1] <-"SME"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(SME_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-SME-DATA.xlsx", row.names = F, showNA = F)

###   18. SNC   ###

#read in company profile data tab and clean to filter rows and columns
SNC_data_global <- read.xlsx("G-SNC.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
SNC_col_clean <- SNC_data_global[,-2]
SNC_row_clean <- SNC_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
SNC_G_no_nas <- SNC_row_clean[rowSums(is.na(SNC_row_clean)) !=ncol(SNC_row_clean), ]
#Change cell A1 to the company TLA
names(SNC_G_no_nas)[1] <-"SNC"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(SNC_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-SNC-DATA.xlsx", row.names = F, showNA = F)

###  19. STA   ###

#read in company profile data tab and clean to filter rows and columns
STA_data_global <- read.xlsx("G-STA.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
STA_col_clean <- STA_data_global[,-2]
STA_row_clean <- STA_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
STA_G_no_nas <- STA_row_clean[rowSums(is.na(STA_row_clean)) !=ncol(STA_row_clean), ]
#Change cell A1 to the company TLA
names(STA_G_no_nas)[1] <-"STA"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(STA_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-STA-DATA.xlsx", row.names = F, showNA = F)

###  20. SWE   ###

#read in company profile data tab and clean to filter rows and columns
SWE_data_global <- read.xlsx("G-SWE.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
SWE_col_clean <- SWE_data_global[,-2]
SWE_row_clean <- SWE_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
SWE_G_no_nas <- SWE_row_clean[rowSums(is.na(SWE_row_clean)) !=ncol(SWE_row_clean), ]
#Change cell A1 to the company TLA
names(SWE_G_no_nas)[1] <-"SWE"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(SWE_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-SWE-DATA.xlsx", row.names = F, showNA = F)

###   21. TAU   ###

#read in company profile data tab and clean to filter rows and columns
TAU_data_global <- read.xlsx("G-TAU.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
TAU_col_clean <- TAU_data_global[,-2]
TAU_row_clean <- TAU_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
TAU_G_no_nas <- TAU_row_clean[rowSums(is.na(TAU_row_clean)) !=ncol(TAU_row_clean), ]
#Change cell A1 to the company TLA
names(TAU_G_no_nas)[1] <-"TAU"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(TAU_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-TAU-DATA.xlsx", row.names = F, showNA = F)

###   22. TET   ###

#read in company profile data tab and clean to filter rows and columns
TET_data_global <- read.xlsx("G-TET.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
TET_col_clean <- TET_data_global[,-2]
TET_row_clean <- TET_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
TET_G_no_nas <- TET_row_clean[rowSums(is.na(TET_row_clean)) !=ncol(TET_row_clean), ]
#Change cell A1 to the company TLA
names(TET_G_no_nas)[1] <-"TET"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(TET_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-TET-DATA.xlsx", row.names = F, showNA = F)

###  23. WOO   ###

#read in company profile data tab and clean to filter rows and columns
WOO_data_global <- read.xlsx("G-WOO.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
WOO_col_clean <- WOO_data_global[,-2]
WOO_row_clean <- WOO_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
WOO_G_no_nas <- WOO_row_clean[rowSums(is.na(WOO_row_clean)) !=ncol(WOO_row_clean), ]
#Change cell A1 to the company TLA
names(WOO_G_no_nas)[1] <-"WOO"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(WOO_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-WOO-DATA.xlsx", row.names = F, showNA = F)

###   24. WOR   ###

#read in company profile data tab and clean to filter rows and columns
WOR_data_global <- read.xlsx("G-WOR.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
WOR_col_clean <- WOR_data_global[,-2]
WOR_row_clean <- WOR_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
WOR_G_no_nas <- WOR_row_clean[rowSums(is.na(WOR_row_clean)) !=ncol(WOR_row_clean), ]
#Change cell A1 to the company TLA
names(WOR_G_no_nas)[1] <-"WOR"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(WOR_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-WOR-DATA.xlsx", row.names = F, showNA = F)

###   25. WSP   ###

#read in company profile data tab and clean to filter rows and columns
WSP_data_global <- read.xlsx("G-WSP.xlsx", sheetIndex = 2, startRow = 8, colIndex = 2:19, as.data.frame = T, check.names = F)
WSP_col_clean <- WSP_data_global[,-2]
WSP_row_clean <- WSP_col_clean[-1,]

#Get rid of rows containing N/A as their row.name
WSP_G_no_nas <- WSP_row_clean[rowSums(is.na(WSP_row_clean)) !=ncol(WSP_row_clean), ]
#Change cell A1 to the company TLA
names(WSP_G_no_nas)[1] <-"WSP"

#Write a new Excel spreadsheet saved in profiles_as_data in the global profiles
write.xlsx(WSP_G_no_nas, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data/G-WSP-DATA.xlsx", row.names = F, showNA = F)

