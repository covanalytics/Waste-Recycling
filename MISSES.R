
#################################
##To Combine all Files if needed#
#################################
#Create a character vector of all EXCEL files
filenames <- list.files(pattern=".xlsx", full.names=T)

#Read the contents of all EXCEl worksheets into a data.frame
df.lists  <- lapply(filenames, function(x) read.xlsx(file=x, sheetIndex=2, startRow=2, colIndex=1:10, as.data.frame=TRUE, header=FALSE))

#Bind the rows of the data.frame lists created from the EXCEL sheets
woMisses  <- rbind.fill(df.lists)

#######################################
#Just read in for current update#######
#######################################
woMisses <- read.xlsx("November2016.xlsx", sheetIndex=2, startRow=2, colIndex=1:10, as.data.frame=TRUE, header=FALSE)

#Add columns names
    names(woMisses)[1] <-"WO_Date"
    names(woMisses)[2] <-"WO_Number"
    names(woMisses)[3] <-"Comp_Num"
    names(woMisses)[4] <-"Cust_Num"
    names(woMisses)[5] <-"PL_Code"
    names(woMisses)[6] <-"WO_Status"
    names(woMisses)[7] <-"Residential"
    names(woMisses)[8] <-"Address"
    names(woMisses)[9] <-"City"
    names(woMisses)[10]<-"State"

#Remove duplicates    
#WO_missedList <- unique(woMisses)

#Create a new column called "Location" and paste content from other columns
#woMisses$Location<- with(woMisses, paste(Address, City, State, sep=", "))

woMisses <- subset(woMisses,WO_Number != 8)    

woMisses <- within(woMisses, {
  ZipCode <- "41011"
  FullAddress <- paste(Address, City, State, ZipCode, sep = " ")})

coordinates <- geocode(woMisses$FullAddress)
woMisses <- cbind(woMisses, coordinates)

### save r ojbect for reloading on next update without having to geocode all records
save(woMisses, file="missesNovember16.RData")
    
###Leave in Address information for geocoding
write.csv(woMisses, "C:/Users/tsink/Mapping/Geocoding/Misses/woMisses_2-26-2016a.csv") 
