# set working directory to where the data is
setwd("~/OneDrive - University Of Houston/Coins Migration")

# load required libraries
require(openxlsx)
require(dplyr)
require(tools)

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
#change variable types
raw_data$Find.Type <- as.factor(raw_data$Find.Type)
raw_data$Find.Territory <- as.factor(raw_data$Find.Territory)
raw_data$Source.Territory <- as.factor(raw_data$Source.Territory)

#rename some columns
names(raw_data)[4] <- "Find.Latitude"
names(raw_data)[5] <- "Find.Longitude"

#counts by territory using the Quantity column
counts_by_territory <- raw_data %>% 
  group_by(Find.Territory, Source.Territory) %>% 
  summarise(counts = sum(Quantity)) 

# show counts_by_territory
print(counts_by_territory)

# check for total coins
print(sum(counts_by_territory$counts))

# write counts to csv file
write.csv(counts_by_territory, file = "counts_by_territory.csv", 
          row.names = FALSE, quote = FALSE)

write.csv(counts_by_region, file= "counts_by_region.csv", row.names = FALSE, quote=FALSE) 