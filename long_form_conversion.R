list.of.packages <- c("xlsx","dplyr")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

setwd("C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data")

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

write.xlsx(AEC_data_reshaped, "C:/Users/adamh/Google Drive/Global MIS/Global profiles/profiles_as_data_long/G-AEC-DATA-LONG.xlsx", row.names = F, showNA = F)
