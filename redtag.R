
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/RedTags/redtag_directory")

library("xlsx")
library("plyr")
library("ggmap")
library("arcgisbinding")
library("sp")
library("spdep")
library("rgdal")
library("maptools")

#Create a character vector of all CSV files
filenames_redtags <- list.files(pattern=".csv", full.names=TRUE)

#Read the contents of all CSV worksheets into a data.frame
redtags.lists  <- lapply(filenames_redtags, function(x) read.csv(file=x, header=FALSE, na.strings = "NA", stringsAsFactors = FALSE, skip=1))

#Bind the rows of the data.frame lists created from the CSV sheets
redtags.bind  <- rbind.fill(redtags.lists)

##Make current month udpate only

#redtags.bind <- read.csv("Red Tag Set Out -1-3-17.csv", header=FALSE, na.strings = "NA", stringsAsFactors = FALSE, skip=1)

#Add columns names
names(redtags.bind) <- c("PID", "Date_Added", "Comp_Num", "Tag_Code", "Cust_Num", "Description","Cust_Name",
                         "St_Num", "St_Name", "St_Suffix", "Service_Zip", "Billing_St", "Billing_St_Name", 
                         "Billing_St_Suffix", "Billing_Zip")

#create file object
redtags <- redtags.bind


#Assign code descriptions to redtags in a new column called Code_Desc
redtags$Code_Desc[redtags$Tag_Code =="CVGAD"] <- "Alternative Disposal"
redtags$Code_Desc[redtags$Tag_Code =="CVGR"]  <- "Billed to Customer"
redtags$Code_Desc[redtags$Tag_Code =="CVGRP"]  <- "Violation Paid by Customer"

#Change code description for billed to customer to indicate those billed to city and those billed to all customers
redtags$Code_Desc[redtags$Code_Desc == "Billed to Customer"] <- "Billed to Customer (City)"

###redtags$Code_Desc[redtags$Code_Desc == "Billed to Customer" & redtags$Cust_Name != "CITY OF COVINGTON"] <- "Billed to Customer (Non-City)"

#Change column to date type
redtags$Date_Added <- as.character(redtags$Date_Added)
redtags$Date_Added <-  as.Date(redtags$Date_Added, format="%Y%m%d")

#Create column called 'Location' and add content to it by pasting strings from other columns
redtags$City <- "Covington"
redtags$State <- "KY"
redtags$Location<- with(redtags, paste(St_Num, St_Name, St_Suffix,City, State, Service_Zip, sep=" "))

#Create data frame with coordinates then bind back to misses data frame
coordinates <- geocode(redtags$Location)
redtags <- cbind(redtags, coordinates)

drops <- names(redtags) %in% c("City", "State")
redtags <- redtags[!drops]


#### Send to ArcGIS ####
send_arcgis <- function(dataframe, path, layerName){
  coordinates(dataframe) <- ~lon + lat
  ## Define Coordinate system for spatial points data.frame 
  reference <- CRS("+init=epsg:4326")
  proj4string(dataframe) <- reference
  ## Assign closest neighborhood and sector in ArcGIS
  writeOGR(obj = dataframe, dsn = path, layer = layerName, driver = 'ESRI Shapefile', overwrite_layer = TRUE)
}
send_arcgis(redtags, "C:/Users/tsink/Mapping/Geocoding/RedTags", "RedTagUpdate")

#### Receive from ArcGIS ####
receive_arcgis <- function(fromPath, dataframeName) {
  arc.check_product()
  ## Read GIS Features 
  read <- arc.open(fromPath)
  ## Create Data.Frame from GIS data 
  dataframeName <- arc.select(read)
  ## Bind hidden lat/long coordinates back to data frame 
  shape <- arc.shape(dataframeName)
  dataframeName<- data.frame(dataframeName, long=shape$x, lat=shape$y)
}
rtagGIS <- receive_arcgis("C:/Users/tsink/Mapping/Geocoding/RedTags/RedTags.shp", rtagGIS)

rtagGIS <- rtagGIS[, c(-1:-2, -20:-23, -25:-35)]
names(rtagGIS) <- c("PID", "Date_Added", "Comp_Num", "Tag_Code", "Cust_Num", "Description","Cust_Name",
                         "St_Num", "St_Name", "St_Suffix", "Service_Zip", "Billing_St", "Billing_St_Name", 
                         "Billing_St_Suffix", "Billing_Zip", "Code_Desc", "Location", "Neighborhood", "lon", "lat")

#### Write Files ####
sql_write <- function(connection, dbName, dbGname, Rfile, dbPull, RepoPath, TablPath, ...) {
  #Database
  connection <- dbConnect(drv=RSQLite::SQLite(), dbname = dbName)
  dbWriteTable(connection, dbGname, Rfile, ...)
  #Pull entire database
  Rfile <- dbGetQuery(connection, dbPull)
  #CovStat Repository
  write.csv(Rfile, file=RepoPath, row.names = FALSE)
  #Tableau File
  write.csv(Rfile, file=TablPath, row.names = FALSE)
  dbDisconnect(connection)
}
connection <- cons.waste
dbName <- "O:/AllUsers/CovStat/Data Portal/repository/Data/Database Files/SolideWaste.db"
dbPull <- 'select * from RedTags'
RepoPath <- "O:/AllUsers/CovStat/Data Portal/Repository/Data/Waste Management/RedTags.csv"
TablPath <- "U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tableau Files/RedTagsTableau.csv"
sql_write(connection, dbName, "RedTags", rtagGIS, dbPull, RepoPath, TablPath, overwrite = TRUE)

