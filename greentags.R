
setwd("U:/Rumpke/GreenTags/PVA")

library("xlsx", lib.loc="~/R/win-library/3.2")
library("plyr", lib.loc="~/R/win-library/3.2")
library("dplyr", lib.loc="~/R/win-library/3.2")
library("tidyr", lib.loc="~/R/win-library/3.2")

#Read the contents of all file into a data.frame
greentags  <-  read.table(file="mosale62(June16).txt", sep = "~", header=FALSE, na.strings = "NA", stringsAsFactors = FALSE)

#Add columns names
names(greentags)[1] <-"OwnerName"
names(greentags)[2] <-"OwnerNameLine2"
names(greentags)[3] <-"OwnerCareofName"
names(greentags)[4] <-"OwnerAddress"
names(greentags)[5] <-"OwnerCity"
names(greentags)[6] <-"OwnerState"
names(greentags)[7] <-"ownerZone"
names(greentags)[8] <-"PIDN"
names(greentags)[9] <-"PropertyLocation"
names(greentags)[10]<-"EditedSaleDate"
names(greentags)[11]<-"EditedBook_Page1"
names(greentags)[12]<-"SalePrice"
names(greentags)[13]<-"SalesCode"
names(greentags)[14]<-"PropertyType"
names(greentags)[15]<-"Landuse_TE"
names(greentags)[16]<-"GreenTag"

#Create column called 'Name' and add content to it by pasting name strings from other columns
greentags$Name<- with(greentags, paste(OwnerName, OwnerNameLine2, sep=""))

#greentags$OwnerName <- NULL
#greentags$OwnerNameLine2 <- NULL
#greentags$OwnerCareofName <- NULL

#Assign green tags "Yes" 
greentags$GreenTag[greentags$Landuse_TE == "SINGLE FAMILY" |  greentags$Landuse_TE == "TWO FAMILY" 
                   | greentags$Landuse_TE == "THREE FAMILY" |  greentags$Landuse_TE == "TOWNHOUSE - NO LAND" 
                   | greentags$Landuse_TE == "CONDOMINIUM" | greentags$Landuse_TE == "MOBILE HOME - LAND" 
                   | greentags$Landuse_TE == "MOBILE HOME - PARK"| greentags$Landuse_TE == "SEASONAL COTTAGE" 
                   | greentags$Landuse_TE == "HOMEOWNERS ASSOCIATION PROPERTY" |  greentags$Landuse_TE == "RES WITH AN OUT BUILDING" 
                   | greentags$Landuse_TE == "LANDOMINIUM" | greentags$Landuse_TE == "FARM LAND W/HOUSE" 
                   | greentags$Landuse_TE == "HORSE FARM W/RESIDENCE" | greentags$Landuse_TE == "DAIRY FARM W/RESIDENCE"
                   | greentags$Landuse_TE == "POULTRY FARM WITH RESIDENCE" | greentags$Landuse_TE == "FRUIT & NUT FARM W/RESIDENCE"
                   | greentags$Landuse_TE == "NURSERY FARM W/RESIDENCE" | greentags$Landuse_TE == "VEGETABLE FARM W/RESIDENCE"
                   | greentags$Landuse_TE == "TOBACCO FARM W/RESIDENCE" | greentags$Landuse_TE == "MIXED FARM W/RESIDENCE"
                   | greentags$Landuse_TE == "FARM LAND W/MOBILE HOME"] <- "Yes"

#Add code descriptions in a column call "CodeDescription" 
greentags$CodeDescription[greentags$SalesCode == "0"] <- "GOOD SALE/ARMS LENGTH TRANSACTION"
greentags$CodeDescription[greentags$SalesCode == "1"] <- "PARTIAL SALE (SPLIT)"
greentags$CodeDescription[greentags$SalesCode == "2"] <- "SALE BETWEEN FAMILY/NOT ARMS LENGTH"
greentags$CodeDescription[greentags$SalesCode == "3"] <- "PROPERTY CLASS CHANGE"
greentags$CodeDescription[greentags$SalesCode == "4"] <- "SALES INCL PERSONAL PROPERTY"
greentags$CodeDescription[greentags$SalesCode == "5"] <- "TRANSCTION INVOLVNG NONTXBL ENTITY"
greentags$CodeDescription[greentags$SalesCode == "6"] <- "TRANSFERRED WITHIN 12 MO"
greentags$CodeDescription[greentags$SalesCode == "7"] <- "SALES LESS THAN $10,000"
greentags$CodeDescription[greentags$SalesCode == "8"] <- "NH SALE/CHANGED BY CONSTRUCTION"
greentags$CodeDescription[greentags$SalesCode == "9"] <- "TRANSFER TAX NOT PAID"
greentags$CodeDescription[greentags$SalesCode == "10"] <- "ADJOINING PROP (ASSEMBLAGE)"
greentags$CodeDescription[greentags$SalesCode == "11"] <- "SALE $ INCL> 1 PROPERTY"
greentags$CodeDescription[greentags$SalesCode == "12"] <- "FORECLOSURE XFER/PURCHASE FROM BANK"
greentags$CodeDescription[greentags$SalesCode == "13"] <- "FULFILLED LAND CONTRACT"
greentags$CodeDescription[greentags$SalesCode == "14"] <- "DELQ. TAX TRANSFER"
greentags$CodeDescription[greentags$SalesCode == "15"] <- "DEED CORRECTION/QUIT CLAIM/MC SALES"
greentags$CodeDescription[greentags$SalesCode == "16"] <- "CORPORATE MERGERS"
greentags$CodeDescription[greentags$SalesCode == "17"] <- "PROPERTY EXCHANGES"
greentags$CodeDescription[greentags$SalesCode == "18"] <- "XFER BETWEEN AFFILIATED COMPANIES"
greentags$CodeDescription[greentags$SalesCode == "19"] <- "SALE PRICE NOT FAIR MARKET VALUE"
greentags$CodeDescription[greentags$SalesCode == "20"] <- "ARMS LENGTH/NEEDS REHAB/ESTATE SALE/HANDYMAN SPECIAL/NOT A FORECLOSURE"
greentags$CodeDescription[greentags$SalesCode == "21"] <- "VACANT LOTS & LAND SALES"
greentags$CodeDescription[greentags$SalesCode == "22"] <- "MOBILE HOME SALES WITH LOTS"
greentags$CodeDescription[greentags$SalesCode == "23"] <- "LAKE PROPERTIES/RIVERFRONT"
greentags$CodeDescription[greentags$SalesCode == "99"] <- "NO CODE ASSIGNED"

#Return only Yes for Green Stickers
greentagsYes <- subset(greentags, GreenTag == "Yes")
greentagsYes <- unique(greentagsYes)

#Export file to EXCEL
write.xlsx(greentagsYes, "U:/Rumpke/GreenTags/PVA/greentags(June).xlsx")


