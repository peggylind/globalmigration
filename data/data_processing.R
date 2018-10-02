# set working directory to where the data is
setwd("~/OneDrive - University Of Houston/Coins Migration/Dataprocessing")

# load required libraries
require(openxlsx)
require(dplyr)
require(tools)
require(tidyr)

# set file name of Excel data file
file <- "Peggy Master List.xlsx"

#read data from Excel
antioch <- read.xlsx(file, 1, rows=c(1:1648))
syria <- read.xlsx(file, 2, rows=c(1:2676))
mesopotamia <- read.xlsx(file, 3, rows=c(1:3067))
southofsyria <- read.xlsx(file, 4, rows=c(1:10001))
turkey <- read.xlsx(file, 5, rows=c(1:5693))
others <- read.xlsx(file, 6, rows=c(1:10488))

#merge all sheets into data frame
raw_data <- rbind(antioch,syria,mesopotamia,southofsyria,turkey,others)

# clean up some data
raw_data$Source.Territory <- trimws(raw_data$Source.Territory)

#change variable types
raw_data$Find.Type <- as.factor(raw_data$Find.Type)
raw_data$Find.Site <- as.factor(raw_data$Find.Site)
raw_data$Find.Territory <- as.factor(raw_data$Find.Territory)
raw_data$Source.Territory <- as.factor(raw_data$Source.Territory)
raw_data$Source.City <- as.factor(raw_data$Source.City)

#print out total number of coins before cleaning
nrow(raw_data)

#rename some columns
names(raw_data)[4] <- "Find.Latitude"
names(raw_data)[5] <- "Find.Longitude"
names(raw_data)[7] <- "Source.Site"

#create chronolocial grouping based on grouping from book (not used right now)
# -350 through -129 = High Hellenistic
# -128 through -31 = Hellenistic to Roman period
# -30 through 95 = Early Roman
# 96 through 192 = High Roman
# 193 through 283 = Severan and Crisis
# 284 through 450 = Late Antiquity
labs <- c("High.Hellenistic", "Hellenistic.to.Roman.period", "Early.Roman",
      "High.Roman", "Severan.and.Crisis", "Late.Antiquity")
first <- cut(raw_data$Fixed.Date.1, breaks=c(-350, -129, -31, 95, 192, 284, 450), include.lowest=TRUE, labels = labs)
second <- cut(raw_data$Fixed.Date.2, breaks=c(-350, -129, -31, 95, 192, 284, 450), include.lowest=TRUE, labels = labs)

#create chronolocial grouping for visualization
#-350 to -64: Hellenistic
#-65 to 192: High Roman
#193 to 284: Late Roman
#294 to 450: Late Antiquity
labs <- c("Hellenistic", "High.Roman",
          "Late.Roman", "Late.Antiquity")
raw_data$period1 = cut(raw_data$Fixed.Date.1, breaks=c(-350, -64, 193, 284, 540), 
                       include.lowest=TRUE, labels = labs)
raw_data$period2 = cut(raw_data$Fixed.Date.2, breaks=c(-350, -64, 193, 284, 540), 
                       include.lowest=TRUE, labels = labs)

# take out rows where period1 or period2 are NA (will take out 1026 cases)
# raw_data <- raw_data[!is.na(raw_data$period1) & !is.na(raw_data$period2),]
# take out rows where period2 are NA (will take out 22 cases)
raw_data <- raw_data[!is.na(raw_data$period2),]

#print out total number of coins after removing rows where Fixed.Data2 == NA
nrow(raw_data)

#take out rows where period1 and period2 are not matching (will take out 2548)
#raw_data <- raw_data[raw_data$period1 == raw_data$period2,]

#create column for chronological period using for counting
raw_data$period <- raw_data$period2

#counts by site using the Quantity column
counts_by_site <- raw_data %>%
  group_by(period, Find.Territory, Source.Territory,Find.Site, Source.Site) %>%
  summarise(sites.counts = sum(Quantity))

#counts by territory using the Quantity column
counts_by_territory <- raw_data %>%
  group_by(period, Find.Territory, Source.Territory) %>%
  summarise(territory.counts = sum(Quantity))

# get regions using unique territories
regions <- unique(unlist(raw_data[,c("Source.Territory", "Find.Territory")]))

# create lookup table for territories/regions
regions.lookup <- data.frame(Territory = as.character(levels(regions)), stringsAsFactors = FALSE)
regions.lookup$ID<-seq.int(nrow(regions.lookup))

# get sites using unique sites
sites <- unique(unlist(raw_data[,c("Source.Site", "Find.Site")]))

# create lookup table sites/countries
sites.lookup <- data.frame(Site = as.character(levels(sites)), stringsAsFactors = FALSE)

# create 3 letter codes for sites/countries
sites.lookup$ISO <- gsub(" ", "", sites.lookup$Site, fixed = TRUE)

# create show from sites where flow counts from one site another are larger than 20
sites_above_20 <- subset(counts_by_site, sites.counts >= 0)
finds <- as.character(sites_above_20$Find.Site)
sources <- as.character(sites_above_20$Source.Site)
uni <- unique(c(finds, sources))
#sites.lookup$show <- ifelse(sites.lookup$Site %in% uni, 1,0)
# create random show=yes/no variable for source and find site/country
#sites.lookup$show <- sample(c(0,1), nrow(sites.lookup), TRUE)
#sites.lookup$show <- 1

# prepare data for D3 migration data structure

#spread the data based on period
counts_by_site <- spread(counts_by_site, period, sites.counts)
counts_by_territory <- spread (counts_by_territory, period, territory.counts)

# get IDs & ISO by match in lookup tables
flow <- counts_by_site
flow$Source.Territory.ID <- with(regions.lookup,
                  ID[match(flow$Source.Territory, Territory)])
flow$Find.Territory.ID <- with(regions.lookup,
                  ID[match(flow$Find.Territory, Territory)])
flow$Source.Site.ISO <- with(sites.lookup,
                  ISO[match(flow$Source.Site, Site)])
flow$Find.Site.ISO <- with(sites.lookup,
                  ISO[match(flow$Find.Site, Site)])
# get region counts from counts_by_territory
flow <- inner_join(flow, counts_by_territory, 
                   by=c("Find.Territory", "Source.Territory"))

#convert NA into 0
flow[is.na(flow)] <- 0

# insert a speace holder column
flow$xxx <- ""

 # reorder columns
flow <- flow[c("Source.Territory.ID", "Source.Territory", "Find.Territory.ID", "Find.Territory",
         "Hellenistic.y", "High.Roman.y", "Late.Roman.y", 
         "Late.Antiquity.y", "xxx", "Source.Site.ISO", "Source.Site", "Find.Site.ISO", "Find.Site",
         "Hellenistic.x", "High.Roman.x", "Late.Roman.x", "Late.Antiquity.x")]

# rename columns
flow <- dplyr::rename(flow, originregion_id=Source.Territory.ID, originregion_name=Source.Territory,
        destinationregion_id=Find.Territory.ID, destinationregion_name=Find.Territory,
        regionflow_1990=Hellenistic.y,  
        regionflow_2095=High.Roman.y, regionflow_2000=Late.Roman.y, regionflow_2005=Late.Antiquity.y, origin_iso=Source.Site.ISO,
        origin_name=Source.Site, destination_iso =Find.Site.ISO, destination_name=Find.Site, 
        countryflow_1990=Hellenistic.x,  
        countryflow_1995=High.Roman.x, countryflow_2000=Late.Roman.x, countryflow_2005=Late.Antiquity.x)

# get show variables from sites lookup table
flow$origin_show <- with(sites.lookup,
                  show[match(flow$origin_name, Site)])
flow$destination_show <- with(sites.lookup,
                 show[match(flow$destination_name, Site)])

# write flow data structure to file
write.csv(flow, file = "coins_flow.csv",
        row.names = FALSE, quote = FALSE)

# write regions to file (to be used in compile.js)
write.table(matrix(as.character(levels(regions)),nrow=1),col.names = FALSE, 
            row.names=FALSE, file = "coins_regions.txt", sep=",")

# write countries.csv using sites.lookup (to be used in filter.js)
sites.lookup <- sites.lookup[c("ISO", "Site", "show")]
sites.lookup <- dplyr::rename(sites.lookup, iso=ISO, name=Site)
write.csv(sites.lookup, file = "coins_country.csv",
          row.names = FALSE, quote = FALSE)