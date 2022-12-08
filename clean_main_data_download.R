#Install packages - Ensure No. TRUE = No. packages
list.of.packages <- c("xlsx","tidyverse","tibble","reshape2")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

#set working directory
setwd("G:/.shortcut-targets-by-id/0B-NghCnC2vhYTU5rOGNaQTBqcHM/Global MIS/2022/Top 100/Adam Top 100")

###CREATE THE 2019/20 DATA FILE###

#read in dataset#
top100_dataset <- read.xlsx("main_data_download.xlsx", sheetIndex = 8, startRow = 2, as.data.frame = T, check.names = F)

#remove col 15 to 28#
top100_dataset_19_20 <- top100_dataset[ ,-15:-28]
#rename our blank column Year#
names(top100_dataset_19_20)[11] <- "Year"
#Assign every row of the Year column the 2019 attribute#
top100_dataset_19_20$Year <- 2019

#Rename our columns to simplify#
names(top100_dataset_19_20)[12] <- "Gross revenue"
names(top100_dataset_19_20)[13] <- "E&S revenue"
names(top100_dataset_19_20)[14] <- "E&S proportion"

#save the file
write.xlsx(top100_dataset_19_20, "top100_19_20.xlsx", row.names = F)




###CREATE THE 2020/21 DATA FILE###

#remove col 12, 13 & 14 then 15 to 25#
top100_dataset_prelim_20_21 <- top100_dataset[ ,-12:-14]
top100_dataset_20_21 <- top100_dataset_prelim_20_21[ ,-15:-25]

#rename our blank column Year#
names(top100_dataset_20_21)[11] <- "Year"
#Assign every row of the Year column the 2020 attribute#
top100_dataset_20_21$Year <- 2020

#Rename our columns to simplify#
names(top100_dataset_20_21)[12] <- "Gross revenue"
names(top100_dataset_20_21)[13] <- "E&S revenue"
names(top100_dataset_20_21)[14] <- "E&S proportion"

#save the file
write.xlsx(top100_dataset_20_21, "top100_20_21.xlsx", row.names = F)



###CREATE THE 2021/22 DATA FILE###

#remove col12 to 18 then 15 to 21#
top100_dataset_prelim_21_22 <- top100_dataset[ ,-12:-18]
top100_dataset_21_22 <- top100_dataset_prelim_21_22[ ,-15:-21]

#rename our blank column Year#
names(top100_dataset_21_22)[11] <- "Year"
#Assign every row of the Year column the 2021 attribute#
top100_dataset_21_22$Year <- 2021

#Rename our columns to simplify#
names(top100_dataset_21_22)[12] <- "Gross revenue"
names(top100_dataset_21_22)[13] <- "E&S revenue"
names(top100_dataset_21_22)[14] <- "E&S proportion"

#save the file
write.xlsx(top100_dataset_21_22, "top100_21_22.xlsx", row.names = F)



###rbind the files to stack them###
top100_bind <- rbind(top100_dataset_19_20, top100_dataset_20_21, top100_dataset_21_22)
#write the file as a new xlsx#
write.xlsx(top100_bind, "top_100_bind.xlsx", row.names = F)
