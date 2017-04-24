
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Misses/misses_directory")

library("xlsx")
library("plyr")
library("ggmap")
library("RSQLite")
library("sp")
library("spdep")
library("rgdal")
library("maptools")
library("readxl")


##To Combine all Files if needed#

#Create a character vector of all EXCEL files
#filenames <- list.files(pattern=".xlsx", full.names=T)

#Read the contents of all EXCEl worksheets into a data.frame
#df.lists  <- lapply(filenames, function(x) read.xlsx(file=x, sheetIndex=2, startRow=2, colIndex=1:10, as.data.frame=TRUE, header=FALSE))

#Bind the rows of the data.frame lists created from the EXCEL sheets
#woMisses  <- rbind.fill(df.lists)

##Load Update
#Be sure to delete first sheet and last row containing summary; some odd behavior with date field if not.  File does not load.
load_data <- function(FileName){
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
woMisses <- load_data("March2017.xlsx")


#### Geocode ####
woMisses <- within(woMisses, {
  ZipCode <- "41011"
  FullAddress <- paste(Address, City, State, ZipCode, sep = " ")})

# Geocode location of permits
covstat_geocode <- function(df, address){
  coordinates <- geocode(address)
  df <- cbind(df, coordinates)
  
  # Add zero to NAs in lat and lon if geocode fails
  if(sum(is.na(df$lat > 0))){
    for (i in 1:length(df$lat)){
      if(is.na(df$lat[i]))
        df$lat[i] <- 0
    }
    for (i in 1:length(df$lon)){
      if(is.na(df$lon[i]))
        df$lon[i] <- 0
    }
  }
  df
}
woMisses <- covstat_geocode(woMisses, woMisses$FullAddress)

# Create SpatialPointsDataFrame
covstat_cr8_SPpts <- function(dataframe, 
                              crs = "+proj=longlat +datum=WGS84", 
                              writeShpFile = FALSE, 
                              path = NULL, 
                              layerName = NULL){
# Set spatial coordinates
  coordinates(dataframe) <- ~lon + lat
  ## Define Coordinate system for spatial points data.frame 
  reference <- CRS(crs)
  proj4string(dataframe) <- reference
  
  # write spatial points data frame as shapefile
  if(writeShpFile){
    writeOGR(obj = dataframe, dsn = path, layer = layerName, driver = 'ESRI Shapefile', overwrite_layer = TRUE)
  } 
  dataframe
}
woMissesSp <- covstat_cr8_SPpts(woMisses)

saveRDS(woMisses, "woMisses.rds")

# Read polygon shapefile and convert to SpatialPolygonsDataFrame, and assign same CRS as points SP
neigh <- readOGR("U:/Mapping/Neighborhoods.shp")
neighPolySp <- spTransform(neigh, CRS("+proj=longlat +datum=WGS84"))



covstat_join_poly2pt <- function(pts, poly, dfOriginal){
  # Assign a polygon's attributes by row position to each point that falls within it
  polyAttrs <- as.data.frame(over(pts, poly))
  # bind polygon attributes to the original, non-sp dataframe.  
  cbind(dfOriginal, polyAttrs)
}
woMisses <- covstat_join_poly2pt(woMissesSp, neighPolySp, woMisses)


woMissesUpdate <- woMisses

############################
##Invalid coordinates and neighborhood indication
woMissesUpdate$NbhdLabel <- as.character(woMissesUpdate$NbhdLabel)
woMissesUpdate$NbhdLabel[which(is.na(woMissesUpdate$NbhdLabel))] <- "Invalid Coordinates"
woMissesUpdate$lat[woMissesUpdate$lon > 0] <- 0
woMissesUpdate$lon[woMissesUpdate$lon > 0] <- 0

### Columns to name and keep
woMissesUpdate <- woMissesUpdate[, -c(12, 15:18, 20:29)]
names(woMissesUpdate) <- c("WO_Date", "WO_Number", "Comp_Num", "Cust_Num", "PL_Code", "WO_Status", "Residentia",
                           "Address", "City", "State", "FullAddres", "lon", "lat", "NbhdLabel")

#Important to have WO_Date class as character to be consistent in db and for correct format to update the tableau dashboard. 
woMissesUpdate$WO_Date <- as.character(woMissesUpdate$WO_Date)

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
dbGname <- "WO_Misses"
dbPull <- 'select * from WO_Misses'
RepoPath <- "O:/AllUsers/CovStat/Data Portal/Repository/Data/Waste Management/WO_Misses.csv"
TablPath <- "U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tableau Files/WO_Misses.csv"

sql_write(connection, dbName , dbGname, Rfile = woMissesUpdate, dbPull, RepoPath, TablPath, append = TRUE)

# Db pull 
cons.waste <- dbConnect(drv=RSQLite::SQLite(), "O:/AllUsers/CovStat/Data Portal/repository/Data/Database Files/SolideWaste.db")
dash_misses <- dbGetQuery(cons.waste, 'select * from WO_Misses')
dash_misses <- dash_misses[1:2716,]



