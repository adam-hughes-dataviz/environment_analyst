list.of.packages <- c("xlsx","dplyr","tidyverse", "readxl","janitor")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

#setting working directory 
wd <- setwd("C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data")

###   AEC   ###

#read in short form data
AEC_Short_data <- read.xlsx("G-AEC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.AEC <- as.data.frame(t(AEC_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.AEC <- tibble::rownames_to_column(t.AEC, "VO")
#rename cell A1 "Year" 
t.AEC[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.AEC.r <- t.AEC %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.AEC.r$company <- "AEC"
t.AEC.r$company_name <- "AECOM"
#convert from character to numeric
i <- c(1:94)
t.AEC.r[ ,i] <- apply(t.AEC.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.AEC.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.AEC.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.AEC.xlsx", row.names = F, showNA = F)

###   ANT   ###

#read in short form data
ANT_Short_data <- read.xlsx("G-ANT-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.ANT <- as.data.frame(t(ANT_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.ANT <- tibble::rownames_to_column(t.ANT, "VO")
#rename cell A1 "Year" 
t.ANT[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.ANT.r <- t.ANT %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.ANT.r$company <- "ANT"
t.ANT.r$company_name <- "Anthesis"
#convert from character to numeric
i <- c(1:94)
t.ANT.r[ ,i] <- apply(t.ANT.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.ANT.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.ANT.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.ANT.xlsx", row.names = F, showNA = F)

###   ARC   ###

#read in short form data
ARC_Short_data <- read.xlsx("G-ARC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.ARC <- as.data.frame(t(ARC_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.ARC <- tibble::rownames_to_column(t.ARC, "VO")
#rename cell A1 "Year" 
t.ARC[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.ARC.r <- t.ARC %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.ARC.r$company <- "ARC"
t.ARC.r$company_name <- "Arcadis"
#convert from character to numeric
i <- c(1:94)
t.ARC.r[ ,i] <- apply(t.ARC.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.ARC.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.ARC.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.ARC.xlsx", row.names = F, showNA = F)

###   CAR   ###

#read in short form data
CAR_Short_data <- read.xlsx("G-CAR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.CAR <- as.data.frame(t(CAR_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.CAR <- tibble::rownames_to_column(t.CAR, "VO")
#rename cell A1 "Year" 
t.CAR[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.CAR.r <- t.CAR %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.CAR.r$company <- "CAR"
t.CAR.r$company_name <- "Cardno"
#convert from character to numeric
i <- c(1:94)
t.CAR.r[ ,i] <- apply(t.CAR.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.CAR.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.CAR.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.CAR.xlsx", row.names = F, showNA = F)

###   ERM   ###

#read in short form data
ERM_Short_data <- read.xlsx("G-ERM-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.ERM <- as.data.frame(t(ERM_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.ERM <- tibble::rownames_to_column(t.ERM, "VO")
#rename cell A1 "Year" 
t.ERM[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.ERM.r <- t.ERM %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.ERM.r$company <- "ERM"
t.ERM.r$company_name <- "ERM"
#convert from character to numeric
i <- c(1:94)
t.ERM.r[ ,i] <- apply(t.ERM.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.ERM.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.ERM.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.ERM.xlsx", row.names = F, showNA = F)

###   GHD   ###

#read in short form data
GHD_Short_data <- read.xlsx("G-GHD-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.GHD <- as.data.frame(t(GHD_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.GHD <- tibble::rownames_to_column(t.GHD, "VO")
#rename cell A1 "Year" 
t.GHD[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.GHD.r <- t.GHD %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.GHD.r$company <- "GHD"
t.GHD.r$company_name <- "GHD"
#convert from character to numeric
i <- c(1:94)
t.GHD.r[ ,i] <- apply(t.GHD.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.GHD.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.GHD.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.GHD.xlsx", row.names = F, showNA = F)

###   GOL   ###

#read in short form data
GOL_Short_data <- read.xlsx("G-GOL-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.GOL <- as.data.frame(t(GOL_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.GOL <- tibble::rownames_to_column(t.GOL, "VO")
#rename cell A1 "Year" 
t.GOL[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.GOL.r <- t.GOL %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.GOL.r$company <- "GOL"
t.GOL.r$company_name <- "Golder"
#convert from character to numeric
i <- c(1:94)
t.GOL.r[ ,i] <- apply(t.GOL.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.GOL.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.GOL.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.GOL.xlsx", row.names = F, showNA = F)

###   HDR   ###

#read in short form data
HDR_Short_data <- read.xlsx("G-HDR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.HDR <- as.data.frame(t(HDR_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.HDR <- tibble::rownames_to_column(t.HDR, "VO")
#rename cell A1 "Year" 
t.HDR[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.HDR.r <- t.HDR %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.HDR.r$company <- "HDR"
t.HDR.r$company_name <- "HDR"
#convert from character to numeric
i <- c(1:94)
t.HDR.r[ ,i] <- apply(t.HDR.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.HDR.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.HDR.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.HDR.xlsx", row.names = F, showNA = F)

###   ICF   ###

#read in short form data
ICF_Short_data <- read.xlsx("G-ICF-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.ICF <- as.data.frame(t(ICF_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.ICF <- tibble::rownames_to_column(t.ICF, "VO")
#rename cell A1 "Year" 
t.ICF[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.ICF.r <- t.ICF %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.ICF.r$company <- "ICF"
t.ICF.r$company_name <- "ICF"
#convert from character to numeric
i <- c(1:94)
t.ICF.r[ ,i] <- apply(t.ICF.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.ICF.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.ICF.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.ICF.xlsx", row.names = F, showNA = F)

###   INO   ###

#read in short form data
INO_Short_data <- read.xlsx("G-INO-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.INO <- as.data.frame(t(INO_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.INO <- tibble::rownames_to_column(t.INO, "VO")
#rename cell A1 "Year" 
t.INO[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.INO.r <- t.INO %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.INO.r$company <- "INO"
t.INO.r$company_name <- "Inogen"
#convert from character to numeric
i <- c(1:94)
t.INO.r[ ,i] <- apply(t.INO.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.INO.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.INO.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.INO.xlsx", row.names = F, showNA = F)

###   JAC   ###

#read in short form data
JAC_Short_data <- read.xlsx("G-JAC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.JAC <- as.data.frame(t(JAC_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.JAC <- tibble::rownames_to_column(t.JAC, "VO")
#rename cell A1 "Year" 
t.JAC[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.JAC.r <- t.JAC %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.JAC.r$company <- "JAC"
t.JAC.r$company_name <- "Jacobs"
#convert from character to numeric
i <- c(1:94)
t.JAC.r[ ,i] <- apply(t.JAC.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.JAC.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.JAC.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.JAC.xlsx", row.names = F, showNA = F)

###   MOT   ###

#read in short form data
MOT_Short_data <- read.xlsx("G-MOT-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.MOT <- as.data.frame(t(MOT_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.MOT <- tibble::rownames_to_column(t.MOT, "VO")
#rename cell A1 "Year" 
t.MOT[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.MOT.r <- t.MOT %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.MOT.r$company <- "MOT"
t.MOT.r$company_name <- "Mott MacDonald"
#convert from character to numeric
i <- c(1:94)
t.MOT.r[ ,i] <- apply(t.MOT.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.MOT.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.MOT.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.MOT.xlsx", row.names = F, showNA = F)

###   RAM   ###

#read in short form data
RAM_Short_data <- read.xlsx("G-RAM-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.RAM <- as.data.frame(t(RAM_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.RAM <- tibble::rownames_to_column(t.RAM, "VO")
#rename cell A1 "Year" 
t.RAM[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.RAM.r <- t.RAM %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.RAM.r$company <- "RAM"
t.RAM.r$company_name <- "Ramboll"
#convert from character to numeric
i <- c(1:94)
t.RAM.r[ ,i] <- apply(t.RAM.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.RAM.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.RAM.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.RAM.xlsx", row.names = F, showNA = F)

###   RHD   ###

#read in short form data
RHD_Short_data <- read.xlsx("G-RHD-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.RHD <- as.data.frame(t(RHD_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.RHD <- tibble::rownames_to_column(t.RHD, "VO")
#rename cell A1 "Year" 
t.RHD[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.RHD.r <- t.RHD %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.RHD.r$company <- "RHD"
t.RHD.r$company_name <- "Royal HaskoningDHV"
#convert from character to numeric
i <- c(1:94)
t.RHD.r[ ,i] <- apply(t.RHD.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.RHD.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.RHD.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.RHD.xlsx", row.names = F, showNA = F)

###   RPS   ###

#read in short form data
RPS_Short_data <- read.xlsx("G-RPS-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.RPS <- as.data.frame(t(RPS_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.RPS <- tibble::rownames_to_column(t.RPS, "VO")
#rename cell A1 "Year" 
t.RPS[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.RPS.r <- t.RPS %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.RPS.r$company <- "RPS"
t.RPS.r$company_name <- "RPS"
#convert from character to numeric
i <- c(1:94)
t.RPS.r[ ,i] <- apply(t.RPS.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.RPS.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.RPS.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.RPS.xlsx", row.names = F, showNA = F)

###   SLR   ###

#read in short form data
SLR_Short_data <- read.xlsx("G-SLR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.SLR <- as.data.frame(t(SLR_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.SLR <- tibble::rownames_to_column(t.SLR, "VO")
#rename cell A1 "Year" 
t.SLR[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.SLR.r <- t.SLR %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.SLR.r$company <- "SLR"
t.SLR.r$company_name <- "SLR Consulting"
#convert from character to numeric
i <- c(1:94)
t.SLR.r[ ,i] <- apply(t.SLR.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.SLR.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.SLR.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.SLR.xlsx", row.names = F, showNA = F)

###   SME   ###

#read in short form data
SME_Short_data <- read.xlsx("G-SME-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.SME <- as.data.frame(t(SME_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.SME <- tibble::rownames_to_column(t.SME, "VO")
#rename cell A1 "Year" 
t.SME[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.SME.r <- t.SME %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.SME.r$company <- "SME"
t.SME.r$company_name <- "Smec Holdings"
#convert from character to numeric
i <- c(1:94)
t.SME.r[ ,i] <- apply(t.SME.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.SME.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.SME.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.SME.xlsx", row.names = F, showNA = F)

###   SNC   ###

#read in short form data
SNC_Short_data <- read.xlsx("G-SNC-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.SNC <- as.data.frame(t(SNC_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.SNC <- tibble::rownames_to_column(t.SNC, "VO")
#rename cell A1 "Year" 
t.SNC[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.SNC.r <- t.SNC %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.SNC.r$company <- "SNC"
t.SNC.r$company_name <- "SNC-Lavalin"
#convert from character to numeric
i <- c(1:94)
t.SNC.r[ ,i] <- apply(t.SNC.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.SNC.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.SNC.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.SNC.xlsx", row.names = F, showNA = F)

###   STA   ###

#read in short form data
STA_Short_data <- read.xlsx("G-STA-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.STA <- as.data.frame(t(STA_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.STA <- tibble::rownames_to_column(t.STA, "VO")
#rename cell A1 "Year" 
t.STA[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.STA.r <- t.STA %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.STA.r$company <- "STA"
t.STA.r$company_name <- "Stantec"
#convert from character to numeric
i <- c(1:94)
t.STA.r[ ,i] <- apply(t.STA.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.STA.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.STA.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.STA.xlsx", row.names = F, showNA = F)

###   SWE   ###

#read in short form data
SWE_Short_data <- read.xlsx("G-SWE-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.SWE <- as.data.frame(t(SWE_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.SWE <- tibble::rownames_to_column(t.SWE, "VO")
#rename cell A1 "Year" 
t.SWE[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.SWE.r <- t.SWE %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.SWE.r$company <- "SWE"
t.SWE.r$company_name <- "Sweco"
#convert from character to numeric
i <- c(1:94)
t.SWE.r[ ,i] <- apply(t.SWE.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.SWE.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.SWE.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.SWE.xlsx", row.names = F, showNA = F)

###   TAU   ###

#read in short form data
TAU_Short_data <- read.xlsx("G-TAU-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.TAU <- as.data.frame(t(TAU_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.TAU <- tibble::rownames_to_column(t.TAU, "VO")
#rename cell A1 "Year" 
t.TAU[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.TAU.r <- t.TAU %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.TAU.r$company <- "TAU"
t.TAU.r$company_name <- "Tauw"
#convert from character to numeric
i <- c(1:94)
t.TAU.r[ ,i] <- apply(t.TAU.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.TAU.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.TAU.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.TAU.xlsx", row.names = F, showNA = F)

###   TET   ###

#read in short form data
TET_Short_data <- read.xlsx("G-TET-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.TET <- as.data.frame(t(TET_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.TET <- tibble::rownames_to_column(t.TET, "VO")
#rename cell A1 "Year" 
t.TET[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.TET.r <- t.TET %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.TET.r$company <- "TET"
t.TET.r$company_name <- "Tetra Tech"
#convert from character to numeric
i <- c(1:94)
t.TET.r[ ,i] <- apply(t.TET.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.TET.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.TET.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.TET.xlsx", row.names = F, showNA = F)

###   WOO   ###

#read in short form data
WOO_Short_data <- read.xlsx("G-WOO-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.WOO <- as.data.frame(t(WOO_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.WOO <- tibble::rownames_to_column(t.WOO, "VO")
#rename cell A1 "Year" 
t.WOO[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.WOO.r <- t.WOO %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.WOO.r$company <- "WOO"
t.WOO.r$company_name <- "Wood Group"
#convert from character to numeric
i <- c(1:94)
t.WOO.r[ ,i] <- apply(t.WOO.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.WOO.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.WOO.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.WOO.xlsx", row.names = F, showNA = F)

###   WOR   ###

#read in short form data
WOR_Short_data <- read.xlsx("G-WOR-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.WOR <- as.data.frame(t(WOR_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.WOR <- tibble::rownames_to_column(t.WOR, "VO")
#rename cell A1 "Year" 
t.WOR[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.WOR.r <- t.WOR %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.WOR.r$company <- "WOR"
t.WOR.r$company_name <- "Advisian"
#convert from character to numeric
i <- c(1:94)
t.WOR.r[ ,i] <- apply(t.WOR.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.WOR.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.WOR.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.WOR.xlsx", row.names = F, showNA = F)

###   WSP   ###

#read in short form data
WSP_Short_data <- read.xlsx("G-WSP-DATA.xlsx", sheetIndex = 1, as.data.frame = T, check.names = F)
#transpose short form data 
t.WSP <- as.data.frame(t(WSP_Short_data), stringsAsFactors = F)
#bring the row names into a column 
t.WSP <- tibble::rownames_to_column(t.WSP, "VO")
#rename cell A1 "Year" 
t.WSP[1,1] = "Year"
#shift first row up to column headings and allow for duplicates, renaming as such 
t.WSP.r <- t.WSP %>%
  row_to_names(row_number = 1) %>%
  clean_names()
#Add a column for organisation name 
t.WSP.r$company <- "WSP"
t.WSP.r$company_name <- "WSP"
#convert from character to numeric
i <- c(1:94)
t.WSP.r[ ,i] <- apply(t.WSP.r[ ,i], 2, 
                      function(x) as.numeric(as.character(x)))
#Add global categorisation
t.WSP.r$global_uk <- "Global"

#write new file in new folder
write.xlsx(t.WSP.r, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/t.WSP.xlsx", row.names = F, showNA = F)

#equate colnames
colnames(t.AEC.r)=colnames(t.ANT.r)=colnames(t.ARC.r)=colnames(t.CAR.r)=colnames(t.ERM.r)=colnames(t.GHD.r)=colnames(t.GOL.r)=colnames(t.HDR.r)=
  colnames(t.ICF.r)=colnames(t.INO.r)=colnames(t.JAC.r)=colnames(t.MOT.r)=colnames(t.RAM.r)=colnames(t.RHD.r)=colnames(t.RPS.r)=colnames(t.SLR.r)=
  colnames(t.SME.r)=colnames(t.SNC.r)=colnames(t.STA.r)=colnames(t.SWE.r)=colnames(t.TAU.r)=colnames(t.TET.r)=colnames(t.WOO.r)=colnames(t.WOR.r)=colnames(t.WSP.r)
#rbind together 
big.bind.global <- rbind(t.AEC.r, t.ANT.r, t.ARC.r, t.CAR.r, t.ERM.r, t.GHD.r, t.GOL.r, t.HDR.r, t.ICF.r, t.INO.r, t.JAC.r, t.MOT.r, t.RAM.r, t.RHD.r,
                         t.RPS.r, t.SLR.r, t.SME.r, t.SNC.r, t.STA.r, t.SWE.r, t.TAU.r, t.TET.r, t.WOO.r, t.WOR.r, t.WSP.r)

#Change column names 
colnames(big.bind.global) <- c("Year", "[Finances] Local - Group", "[Finances] Local - Environmental services", "[Finances] Local - Environmental consultancy",
                       "[Finances] CER US$ - Group", "[Finances] CER US$ - Environmental services", "[Finances] CER US$ - Environmental consultancy",
                       "[Finances] Override US$ - Group", "[Finances] Override US$ - Environmental services", "[Finances] Override US$ - Environmental consultancy",
                       "[Finances] Difference US$ - Group", "[Finances] Difference US$ - Environmental services", "[Finances] Difference US$ - Environmental consultancy",
                       "[Finances] Used US$ - Group", "[Finances] Used US$ - Environmental services","[Finances] Used US$ - Environmental consultancy",
                       "[Finances] Other group", "[Finances] Other environmental services", "[Finances] Environmental consultancy (EC)",
                       "[Finances] Local growth - Group", "[Finances] Local growth - Environmental services", "[Finances] Local growth - Environmental consultancy",
                       "[Finances] Used US$ growth - Group", "[Finances] Used US$ growth - Environmental services", "[Finances] Used US$ growth - Environmental consultancy",
                       "[Finances] % of work that is EC", "[Finances] EC growth", "[Finances] 3 year growth - Group", "[Finances] 3 year growth - Environmental services", "[Finances] 3 year growth - Environmental consultancy",
                       "[Finances] Operating profit (EBITDA) - Group", "[Finances] Operating margin - Group (% of gross turnover",
                       "[Staff] Staff - Group (average for year)", "[Staff] Staff - Environmental services", "[Staff] Staff - EC",
                       "[Staff] Turnover per head - Environmental", "[Staff] Turnover per head - EC",
                       "[Staff] Western Europe", "[Staff] E Europe/FSU", "[Staff] Africa and Middle East", "[Staff] Asia-Pacific", "[Staff] North America", "[Staff] Latin America", "[Staff] x-check",
                       "[Contracts] Contracts", "[Contracts] Average contract value",
                       "[Service areas] Climate change & energy, %","[Service areas] Contaminated site assessment/remediation, %", "[Service areas] Impact assessment, %", "[Service areas] Env. due diligence, management and compliance, %", "[Service areas] Water & waste management, %", "[Service areas] Other services, %",
                       "[Service areas] Climate change & energy","[Service areas] Contaminated site assessment/remediation", "[Service areas] Impact assessment", "[Service areas] Env. due diligence, management and compliance", "[Service areas] Water & waste management", "[Service areas] Other services",
                       "[Clients] Government and regulators, %", "[Clients] Energy anbd utilities (inc. waste), %", "[Clients] Mining, manufacturing & process industries, %", "[Clients] Infrastructure & development, %", "[Clients] Financial, professional & service sector, %", "[Clients] Other, %",
                       "[Clients] Government and regulators", "[Clients] Energy anbd utilities (inc. waste)", "[Clients] Mining, manufacturing & process industries", "[Clients] Infrastructure & development", "[Clients] Financial, professional & service sector", "[Clients] Other",
                       "[Locations] Western Europe, %", "[Locations] E Europe/FSU, %", "[Locations] Africa and Middle East, %", "[Locations] Asia-Pacific, %", "[Locations] North America, %", "[Locations] L-America, %", "[Locations] X-check - Total",
                       "[Locations] Western Europe (% EC)", "[Locations] E Europe/FSU (% EC)", "[Locations] Africa and Middle East (% EC)", "[Locations] Asia-Pacific (% EC)", "[Locations] North America (%EC)", "[Locations] L-America (%EC)",
                       "[Locations] Western Europe", "[Locations]  E Europe/FSU","[Locations] Africa and Middle East", "[Locations] Asia-Pacific", "[Locations] North America", "[Locations] Latin America", "[Locations] Total",
                       "XYX Consultancy", "[US] Organic revenue", "[US] Organic growth", "[US] Organic growth, %", "Company TLA", "Company name", "Global/UK")
#write combined file
write.xlsx(big.bind.global, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/big.bind.global.xlsx", row.names = F, showNA = F)


