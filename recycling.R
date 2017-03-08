
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Recycling/recycling_directory")

options(java.parameters = "-Xmx1000m")

library("xlsx")
library("plyr")
library("dplyr")
library("tidyr")
library("reshape")
library("reshape2")
library("stringr")
library("zoo")
library("arcgisbinding")
library("sp")
library("spdep")
library("rgdal")
library("maptools")
library("ggmap")

#Read the contents of the Recycling worksheet into a data.frame


recyclers  <- read.xlsx2("Recycling-2-1-17.xlsx", sheetName="Recyclers", startRow=3, colIndex=1:13, as.data.frame=TRUE, header=FALSE)

names(recyclers) <- c("Delivery_Date","PID_Num","WO_Route","St_Num","St_Name","St_Suffix","Zip",
                      "St_Direction","Serial.Number","Chrg_Code","Size_Code","Route_Freq","Pickup_Count")

recyclers$City <- "Covington"
recyclers$State <- "KY"
recyclers$recyclers <- "YES"
recyclers <- recyclers[!is.na(recyclers$Delivery_Date),]

#Convert Delivery Date from Numeric to Character format and then to Date format
recyclers$Delivery_Date <- as.character(recyclers$Delivery_Date)
recyclers$Delivery_Date <-  as.Date(recyclers$Delivery_Date, format="%Y%m%d")

#Correct st suffixs for geocoding
recyclers$St_Suffix <- gsub("PIKE", "PI", recyclers$St_Suffix)

#Create a new column called "Location" and paste content from other columns
recyclers$Location<- with(recyclers, paste(St_Num, St_Direction, St_Name, St_Suffix, City, State, Zip, sep=" "))

#Strip out special character from PID Number
#recyclers$PID_Num <-  str_replace_all(recyclers$Location, "[.]", "")

#remove duplicate rows
recyclers <- unique(recyclers)


########################################################################################################

#Read the contents of the Non-recycling worksheet into a data.frame
n_recyclers  <- read.xlsx2("Recycling-2-1-17.xlsx", sheetName="Non Recyclers", startRow=3, colIndex=1:13, as.data.frame=TRUE, header=FALSE)

names(n_recyclers) <- c("Delivery_Date","PID_Num","WO_Route","St_Num","St_Name","St_Suffix","Zip",
                      "St_Direction","Serial.Number","Chrg_Code","Size_Code","Route_Freq","Pickup_Count")

n_recyclers$City <- "Covington"
n_recyclers$State <- "KY"
n_recyclers$recyclers <- "NO"
n_recyclers <- n_recyclers[!is.na(n_recyclers$Delivery_Date),]

#Convert Delivery Date from Numeric to Character format and then to Date format
n_recyclers$Delivery_Date <- as.character(n_recyclers$Delivery_Date)
n_recyclers$Delivery_Date <-  as.Date(n_recyclers$Delivery_Date, format="%Y%m%d")

#Correct st suffixs for geocoding
n_recyclers$St_Suffix <- gsub("PIKE", "PI", n_recyclers$St_Suffix)

#Create a new column called "Location" and paste content from other columns
n_recyclers$Location<- with(n_recyclers, paste(St_Num, St_Direction, St_Name, St_Suffix,  City, State, Zip, sep=" "))

#Strip out special character from PID Number
#n_recyclers$PID_Num <-  str_replace_all(n_recyclers$Location, "[.]", "")

#remove duplicate rows
n_recyclers <- unique(n_recyclers)

###Combine recyclers and non-recyclers data.frames
##Bind rows of Recyclers and Non-recyclers data.frames
recycling <- rbind(recyclers, n_recyclers)


#Write an EXCEL sheet with combined recycling data and coordinate data
write.csv(recycling, "C:/Users/tsink/Mapping/Geocoding/Recycling/recyclingCombined.csv")


############################  STOP HERE ########################################################

#####################
##Connect to ArcGIS##
#####################

receive_arcgis <- function(fromPath, dataframeName) {
  arc.check_product()
  ## Read GIS Features 
  read <- arc.open(fromPath)
  ## Create Data.Frame from GIS data 
  dataframeName <- arc.select(read)
  ## Bind hidden lat/long coordinates back to data frame 
  shape <- arc.shape(dataframeName)
  dataframeName<- data.frame(dataframeName, lon=shape$x, lat=shape$y)
}
recycling_update <- receive_arcgis("C:/Users/tsink/Mapping/Geocoding/Recycling/Recycling(Neighborhoods).shp", 
                                   recycling_update)

##Columns to keep and renames
recycling_update <- recycling_update[, c(16:32, 37)]
#recycling_update <- plyr::rename(recycling_update, c("recyclers"="Recycling Account", "NbhdLabel"="Neighborhood"))

#Remove last '.' in PIDN to match shapefile
#rec <- recycling
recycling_update$PID_Num <- gsub("\\.", "\\-", recycling_update$PID_Num)
substring(recycling_update$PID_Num, 14, 14) <- "."

## Get Parcel shapefile
#inputFC <- "M:/Parcel.shp"
#parcel <- arc.open(inputFC)
#pidn_file <- arc.select(parcel)


## Merge recycling with parcel id file
#rec_trial <- merge(recycling_update, pidn_file, by.x="PID_Num", by.y="PIDN", all.x=TRUE)
options(java.parameters = "-Xmx1000m")
library("RSQLite")
sql_write_recycling <- function(connection, dbName, dbGname, Rfile, RepoPath, ...) {
  #Database
  connection <- dbConnect(drv=RSQLite::SQLite(), dbname = dbName)
  dbWriteTable(connection, dbGname, Rfile, ...)
  #CovStat Repository
  write.csv(Rfile, file=RepoPath, row.names = FALSE)
  dbDisconnect(connection)
}

connection <- cons.waste
dbName <- "O:/AllUsers/CovStat/Data Portal/repository/Data/Database Files/SolideWaste.db"
dbGname <- "Recycling"
Rfile <- recycling_update
RepoPath <- "O:/AllUsers/CovStat/Data Portal/Repository/Data/Waste Management/Recycling.csv"


sql_write_recycling(connection, dbName , dbGname, Rfile, RepoPath, overwrite = TRUE)

##Write to Tableau Dashboard File ##
write.xlsx2(recycling_update, 
            "U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tableau Files/Recycling_With_Attached_Neighborhoods.xlsx"
            ,sheetName = "Recycling_With_Attached_Neighbo", row.names = FALSE)



#### Spatial Data Object ####
#spObject <- arc.data2sp(recycling_arcgis)

#### plot Spatial Data ####
#spplot(spObject)


#### Export Spatial Data to Feature Class ####
#outputFC <- "C:/Users/tsink/Mapping/Geocoding/Recycling/Recycling(Neighborhoods)R.shp"
#arc.write(outputFC, recycling_arcgis)




