
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Misses/misses_directory")

library("xlsx")
library("plyr")
library("ggmap")
library("RSQLite")
library("sp")
library("spdep")
library("rgdal")
library("maptools")

##To Combine all Files if needed#

#Create a character vector of all EXCEL files
#filenames <- list.files(pattern=".xlsx", full.names=T)

#Read the contents of all EXCEl worksheets into a data.frame
#df.lists  <- lapply(filenames, function(x) read.xlsx(file=x, sheetIndex=2, startRow=2, colIndex=1:10, as.data.frame=TRUE, header=FALSE))

#Bind the rows of the data.frame lists created from the EXCEL sheets
#woMisses  <- rbind.fill(df.lists)

##Load Update
#Be sure to delete first sheet; some odd behavior if not.  File does not load.
load_data <- function(df, FileName){
  df <- read.xlsx(FileName, sheetIndex=2, startRow=2, colIndex = 1:12,  as.data.frame=TRUE, header=FALSE)
  #Remove the service type and duplicated customer number field
  df <- df[,c(-6:-7)]
  #Add columns names
  names(df) <- c("WO_Date","WO_Number","Comp_Num","Cust_Num","PL_Code","WO_Status",
                       "Residential","Address","City","State")
  # Excluding certain work orders
  # Can't remember what this is..
  df <- subset(df, WO_Number != 8)
  return(df)
}
woMisses <- load_data(woMisses, "January2017.xlsx")

#### Geocode ####
woMisses <- within(woMisses, {
  ZipCode <- "41011"
  FullAddress <- paste(Address, City, State, ZipCode, sep = " ")})

#Create data frame with coordinates then bind back to misses data frame
coordinates <- geocode(woMisses$FullAddress)
woMisses <- cbind(woMisses, coordinates)

####Add zero to NAs in lat and lon after geocoding----
for (i in 1:length(woMisses$lat)){
  if(is.na(woMisses$lat[i]))
      woMisses$lat[i] <- 0
  }
for (i in 1:length(woMisses$lon)){
  if(is.na(woMisses$lon[i]))
      woMisses$lon[i] <- 0
}

#### Send to ArcGIS ####
send_arcgis <- function(dataframe, path, layerName){
  coordinates(dataframe) <- ~lon+lat
  ## Define Coordinate system for spatial points data.frame 
  reference <- CRS("+init=epsg:4326")
  proj4string(dataframe) <- reference
  ## Assign closest neighborhood and sector in ArcGIS
  writeOGR(obj = dataframe, dsn = path, layer = layerName, driver = 'ESRI Shapefile', overwrite_layer = TRUE)
}
send_arcgis(woMisses, "C:/Users/tsink/Mapping/Geocoding/Misses", "MissesUpdate")
stop('Open in ArcGIS')

#### Receive from ArcGIS ####
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
woMissesUpdate <- receive_arcgis("C:/Users/tsink/Mapping/Geocoding/Misses/MissesUpdateOutput.shp", woMissesUpdate)

### Columns to name and keep
woMissesUpdate <- woMissesUpdate[, c(-1:-2, -14:-18, -20:-30)]
woMissesUpdate <- woMissesUpdate[, c(1:11, 13:14, 12)]
names(woMissesUpdate) <- c("WO_Date", "WO_Number", "Comp_Num", "Cust_Num", "PL_Code", "WO_Status", "Residentia",
                           "Address", "City", "State", "FullAddres", "lon", "lat", "NbhdLabel")

#### Write Files ####
sql_write <- function(connection, dbName, dbGname, Rfile, dbPull, RepoPath, TablPath, ...) {
  #Database
  connection <- dbConnect(drv=RSQLite::SQLite(), dbname = dbName, row.names = FALSE)
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
dbGname <- "WO_Misses"
Rfile <- woMissesUpdate
dbPull <- 'select * from WO_Misses'
RepoPath <- "O:/AllUsers/CovStat/Data Portal/Repository/Data/Waste Management/WO_Misses.csv"
TablPath <- "U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tableau Files/WO_Misses.csv"

sql_write(connection, dbName , dbGname, Rfile, dbPull, RepoPath, TablPath, append = TRUE)




