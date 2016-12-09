
#Read the contents of all file into a data.frame
greentags  <-  read.xlsx2(file="greentags(October).xlsx", sheetName="Covington Query" ,as.data.frame=TRUE, header=TRUE)

#Assign green tags "Yes" 
assign_yes <- function(a = greentags$GreenTag, b =greentags$LANDUSE_TE){
a[b == "SINGLE FAMILY" |  b == "TWO FAMILY" | b == "THREE FAMILY" |  b == "TOWNHOUSE - NO LAND" | 
  b == "CONDOMINIUM" | b == "MOBILE HOME - LAND" | b == "MOBILE HOME - PARK"| b == "SEASONAL COTTAGE" |
  b == "HOMEOWNERS ASSOCIATION PROPERTY" |  b == "RES WITH AN OUT BUILDING" | b == "LANDOMINIUM" | b == "FARM LAND W/HOUSE" |
  b == "HORSE FARM W/RESIDENCE" | b == "DAIRY FARM W/RESIDENCE"| b == "POULTRY FARM WITH RESIDENCE" | b == "FRUIT & NUT FARM W/RESIDENCE"| 
  b == "NURSERY FARM W/RESIDENCE" | b == "VEGETABLE FARM W/RESIDENCE"| b == "TOBACCO FARM W/RESIDENCE" |
  b == "MIXED FARM W/RESIDENCE"| b == "FARM LAND W/MOBILE HOME" | b == "FOUR FAMILY"] <- "Yes"

return(a)
}
greentags$GreenTag <- assign_yes()

greentags$GreenTag[greentags$LANDUSE_TE == "4-19 UNIT APARTMENTS"] <- "Yes(4-19 Unit)"

#Add code descriptions in a column call "CodeDescription" 
add_description <- function(a = greentags$CodeDescription, b =greentags$SALES_CODE){
a[b == "0000"] <- "GOOD SALE/ARMS LENGTH TRANSACTION"
a[b == "0001"] <- "PARTIAL SALE (SPLIT)"
a[b == "0002"] <- "SALE BETWEEN FAMILY/NOT ARMS LENGTH"
a[b == "0003"] <- "PROPERTY CLASS CHANGE"
a[b == "0004"] <- "SALES INCL PERSONAL PROPERTY"
a[b == "0005"] <- "TRANSCTION INVOLVNG NONTXBL ENTITY"
a[b == "0006"] <- "TRANSFERRED WITHIN 12 MO"
a[b == "0007"] <- "SALES LESS THAN $10,000"
a[b == "0008"] <- "NH SALE/CHANGED BY CONSTRUCTION"
a[b == "0009"] <- "TRANSFER TAX NOT PAID"
a[b == "0010"] <- "ADJOINING PROP (ASSEMBLAGE)"
a[b == "0011"] <- "SALE $ INCL> 1 PROPERTY"
a[b == "0012"] <- "FORECLOSURE XFER/PURCHASE FROM BANK"
a[b == "0013"] <- "FULFILLED LAND CONTRACT"
a[b == "0014"] <- "DELQ. TAX TRANSFER"
a[b == "0015"] <- "DEED CORRECTION/QUIT CLAIM/MC SALES"
a[b == "0016"] <- "CORPORATE MERGERS"
a[b == "0017"] <- "PROPERTY EXCHANGES"
a[b == "0018"] <- "XFER BETWEEN AFFILIATED COMPANIES"
a[b == "0019"] <- "SALE PRICE NOT FAIR MARKET VALUE"
a[b == "0020"] <- "ARMS LENGTH/NEEDS REHAB/ESTATE SALE/HANDYMAN SPECIAL/NOT A FORECLOSURE"
a[b == "0021"] <- "VACANT LOTS & LAND SALES"
a[b == "0022"] <- "MOBILE HOME SALES WITH LOTS"
a[b == "0023"] <- "LAKE PROPERTIES/RIVERFRONT"
a[b == "9999"] <- "NO CODE ASSIGNED"

return(a)
}
greentags$CodeDescription <- add_description()


#Return only Yes for Green Stickers
greentagsYes <- subset(greentags, GreenTag == "Yes" | GreenTag == "Yes(4-19 Unit)")
greentagsYes <- unique(greentagsYes)

#Export file to EXCEL
write.xlsx(greentagsYes, "greentags(October).xlsx", row.names = FALSE)

