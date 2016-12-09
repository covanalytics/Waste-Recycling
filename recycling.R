
setwd("U:/Rumpke/Recycling/recycling_directory")

library("xlsx")
library("plyr")

#Read the contents of the Recycling worksheet into a data.frame
recyclers  <- read.xlsx2("Recycling-12-1-16.xlsx", sheetName="Recyclers", startRow=3, colIndex=1:13, as.data.frame=TRUE, header=FALSE)

#Add columns names
names(recyclers)[1] <-"Delivery_Date"
names(recyclers)[2] <-"PID_Num"
names(recyclers)[3] <-"WO_Route"
names(recyclers)[4] <-"St_Num"
names(recyclers)[5] <-"St_Name"
names(recyclers)[6] <-"St_Suffix"
names(recyclers)[7] <-"Zip"
names(recyclers)[8] <-"St_Direction"
names(recyclers)[9] <- "Serial.Number"
names(recyclers)[10] <-"Chrg_Code"
names(recyclers)[11]<-"Size_Code"
names(recyclers)[12]<-"Route_Freq"
names(recyclers)[13]<-"Pickup_Count"

recyclers$City <- "Covington"
recyclers$State <- "KY"
recyclers$recyclers <- "YES"

#Convert Delivery Date from Numeric to Character format and then to Date format
recyclers$Delivery_Date <- as.character(recyclers$Delivery_Date)
recyclers$Delivery_Date <-  as.Date(recyclers$Delivery_Date, format="%Y%m%d")

#Correct st suffixs for geocoding
recyclers$St_Suffix <- gsub("PIKE", "PI", recyclers$St_Suffix)

#Create a new column called "Location" and paste content from other columns
recyclers$Location<- with(recyclers, paste(St_Num, St_Direction, St_Name, St_Suffix, sep=" "))

#Strip out special character from PID Number
#recyclers$PID_Num <-  str_replace_all(recyclers$Location, "[.]", "")

#remove duplicate rows
recyclers <- unique(recyclers)

#Write recyclers to Rfile
save(recyclers, file="recyclers.RData")

#Create an EXCEL file for recyclers
write.xlsx(recyclers, "U:/Rumpke/Recycling/CombinedFiles/recyclers.xlsx")

########################################################################################################

#Read the contents of the Non-recycling worksheet into a data.frame
n_recyclers  <- read.xlsx2("Recycling-12-1-16.xlsx", sheetName="Non-recyclers", startRow=3, colIndex=1:13, as.data.frame=TRUE, header=FALSE)

#Add columns names
names(n_recyclers)[1] <-"Delivery_Date"
names(n_recyclers)[2] <-"PID_Num"
names(n_recyclers)[3] <-"WO_Route"
names(n_recyclers)[4] <-"St_Num"
names(n_recyclers)[5] <-"St_Name"
names(n_recyclers)[6] <-"St_Suffix"
names(n_recyclers)[7] <-"Zip"
names(n_recyclers)[8] <-"St_Direction"
names(n_recyclers)[9] <- "Serial.Number"
names(n_recyclers)[10] <-"Chrg_Code"
names(n_recyclers)[11]<-"Size_Code"
names(n_recyclers)[12]<-"Route_Freq"
names(n_recyclers)[13]<-"Pickup_Count"

n_recyclers$City <- "Covington"
n_recyclers$State <- "KY"
n_recyclers$recyclers <- "NO"

#Convert Delivery Date from Numeric to Character format and then to Date format
n_recyclers$Delivery_Date <- as.character(n_recyclers$Delivery_Date)
n_recyclers$Delivery_Date <-  as.Date(n_recyclers$Delivery_Date, format="%Y%m%d")

#Correct st suffixs for geocoding
n_recyclers$St_Suffix <- gsub("PIKE", "PI", n_recyclers$St_Suffix)

#Create a new column called "Location" and paste content from other columns
n_recyclers$Location<- with(n_recyclers, paste(St_Num, St_Direction, St_Name, St_Suffix,  sep=" "))

#Strip out special character from PID Number
#n_recyclers$PID_Num <-  str_replace_all(n_recyclers$Location, "[.]", "")

#remove duplicate rows
n_recyclers <- unique(n_recyclers)

#Write non-recyclers to Rfile
save(n_recyclers, file="n_recyclers.RData")

#Create an EXCEl file for non-recyclers
write.xlsx(n_recyclers, "U:/Rumpke/Recycling/CombinedFiles/n_recyclers.xlsx")

########################################################################################################
###Combine recyclers and non-recyclers data.frames
##Bind rows of Recyclers and Non-recyclers data.frames
recycling <- rbind(recyclers, n_recyclers)

#Write an EXCEL sheet with combined recycling data and coordinate data
write.csv(recycling, "C:/Users/tsink/Mapping/Geocoding/Recycling/recyclingCombined.csv")

