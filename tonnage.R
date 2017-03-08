
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tonnage/tonnageDirectory")

library("xlsx")
library("plyr")


load_tonnage <- function(df, Filename, sheetName, colIndex){
  df <- read.xlsx(Filename, sheetName=sheetName,  stringsAsFactors=FALSE, colIndex = colIndex, as.data.frame=TRUE)
       if (c("Landfill.Description") %in% names(df)){
          df$Landfill.Description <- as.character(df$Landfill.Description)
          df$SubDistrictDept <- "BUTLER"
          df <- df[, c(6, 1:5)]
          df <- plyr::rename(df, c("SoftPak.Company"="SoftPakCompanyID", "Route.Type.Description"="LOB",
                                   "Landfill.Description"="LandfillDescription", "Fiscal.Date"="Date",
                                   "Tons"="ActTons"))
          message("there")
          return(df)
       }else{
          df$LandfillDescription <- "RUMPKE BUTLER LANDFILL"
          df <- df[, c(1:3, 6, 4:5)]
          df <- plyr::rename(df, c("Sub.District.Name"="SubDistrictDept", "Company.Number"="SoftPakCompanyID", 
                             "Fiscal.Date"="Date", "Tons"="ActTons"))
          message("not there")
          return(df)
       }
  
}
##Front Load Tonnage
tonnageFL <- load_tonnage(tonnageFL, "CovingtonTonnage-Jan2017.xlsx", 'Front Load', c(1:3, 7, 15))
##Rear Load Tonnage
tonnageRL <- load_tonnage(tonnageRL, "CovingtonTonnage-Jan2017.xlsx", 'Rear Load', c(1:3, 7, 15))
##Roll Off
tonnageRO <- load_tonnage(tonnageRO, "CovingtonTonnage-Jan2017.xlsx", 'Roll Off', c(1, 3, 7, 10, 22))
##Put files together
tonnageCombinedFinal <- do.call("rbind", list(tonnageFL, tonnageRL, tonnageRO))

##Last bit of cleaning for waste tonnage
#tonnageCombinedFinal$LandfillDescription[which(is.na(tonnageCombinedFinal$LandfillDescription))]  <- "RUMPKE BUTLER LANDFILL"
#tonnageCombinedFinal$LandfillDescription[tonnageCombinedFinal$LandfillDescription == "Rumpke Butler Landfill"] <- "RUMPKE BUTLER LANDFILL"


##Change directory to get tonnage recycling data
setwd("U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tonnage/tonnageDirectory/Recycling")

###
####Read Tonnage Recycling Data and Bind to Waste Tonnage Data before Writing the file####
##Front Load (FL) Recycling

df.listsR  <- read.xlsx("Jan2017.xlsx", sheetName="Front Load", stringsAsFactors=FALSE, colIndex=1:17, as.data.frame=TRUE)
rec_FL_combined  <- rbind.fill(df.listsR)
keeps <- names(rec_FL_combined) %in% c("Sub.District.Name", "Company.Number", "LOB", "Code.Description", "Fiscal.Date", "Tons")
rec_FL_combined <- rec_FL_combined[keeps]
names(rec_FL_combined) <- c("SubDistrictDept", "SoftPakCompanyID","LOB", "LandfillDescription", "Date", "ActTons")
rec_FL_combined$LandfillDescription <- "Landfill Diversion"

##Rear Load (RL) Recycling

df.listsR  <- read.xlsx("Jan2017.xlsx", sheetName="Rear Load", stringsAsFactors=FALSE, colIndex=1:17, as.data.frame=TRUE)
rec_RL_combined  <- rbind.fill(df.listsR)
keeps <- names(rec_RL_combined) %in% c("Sub.District.Name", "Company.Number", "LOB", "Code.Description", "Fiscal.Date", "Tons")
rec_RL_combined <- rec_RL_combined[keeps]
names(rec_RL_combined) <- c("SubDistrictDept", "SoftPakCompanyID","LOB", "LandfillDescription", "Date", "ActTons")
rec_RL_combined$LandfillDescription <- "Landfill Diversion"

###Put Recycling FL and RL together
recycling_combined <- do.call("rbind", list(rec_FL_combined, rec_RL_combined))

##Bind Waste Tonnage and Recycling Tonnage
combined_WR <- do.call("rbind", list(tonnageCombinedFinal, recycling_combined))
combined_WR$Date <- as.character(combined_WR$Date)


#### Write Files ####
library("RSQLite")
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
dbGname <- "WasteTonnage"
Rfile <- combined_WR
dbPull <- 'select * from WasteTonnage'
RepoPath <- "O:/AllUsers/CovStat/Data Portal/Repository/Data/Waste Management/tonnage.csv"
TablPath <- "U:/CityWide Performance/CovStat/CovStat Projects/Operations/Rumpke/Tableau Files/tonnage.csv"

sql_write(connection, dbName, dbGname, Rfile, dbPull, RepoPath, TablPath, append = TRUE)










