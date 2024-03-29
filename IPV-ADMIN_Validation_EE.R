library(readxl)
library(stringr)

# This separate file is needed because the Estonian Excel questionnaire has located indicators 10 - 13
# two rows down from where they are located in the template.

# Provide here the name of the Excel file you want to validate

fileName <- "Copy of EIGE_Data collection tool_EE.xlsx"

# Provide the filename of the text file in which the findings of the validation will be stored

valFileName <- "EE automated validation report_20231202.txt"

# Remove earlier versions of the file. A warning message will appear if no such file exists but it is normal and can be ignored.
# ATTENTION: if you want to retain earlier versions, keep copies in a different folder before executing the code

file.remove(valFileName)

# Create 'shell' frames that give the 'coordinates' of each cell: the indicator, type of relationship, victim's sex and year

cellIndicator <- data.frame(matrix("",nrow=84,ncol=9))
cellRelationship <- data.frame(matrix("",nrow=84,ncol=9))
cellSex <- data.frame(matrix("",nrow=84,ncol=9))
cellYear <- data.frame(matrix(0,nrow=84,ncol=9))

cellIndicator[1:6,] <- "Indicator 1: Annual number of victims of violence, as recorded by the police"
cellIndicator[7:12,] <- "Indicator 2: Annual number of reported offences of violence against victims, as recorded by the police"
cellIndicator[13:18,] <- "Indicator 3: Annual number of male perpetrators of violence against victims, as recorded by the police"
cellIndicator[19:24,] <- "Indicator 4: Annual number of victims of physical violence, as recorded by the police"
cellIndicator[25:30,] <- "Indicator 5: Annual number of victims of psychological violence, as recorded by the police"
cellIndicator[31:36,] <- "Indicator 6: Annual number of victims of sexual violence, as recorded by the police"
cellIndicator[37:42,] <- "Indicator 7: Annual number of victims of economic violence, as recorded by the police"
cellIndicator[43:48,] <- "Indicator 8: Annual number of victims reporting rape, as recorded by the police"
cellIndicator[49:54,] <- "Indicator 9: Annual number of victims of [femicide/homicide], as recorded by the police"
cellIndicator[55:66,] <- "Indicator 10: Annual number of protection orders in cases of violence against victims"
cellIndicator[67:72,] <- "Indicator 11: Annual number of male perpetrators prosecuted for violence against victims"
cellIndicator[73:78,] <- "Indicator 12: Annual number of male perpetrators sentenced for violence against victims"
cellIndicator[79:84,] <- "Indicator 13: Number of male perpetrators held in prison or with a sanction involving a form of deprivation of liberty for violence against victims (on Dec. 31st of each year)"

cellRelationship[1:2,] <- "Intimate partner relationships"
cellRelationship[3:4,] <- "Domestic relationships"
cellRelationship[5:6,] <- "Any relationship (known or unknown)"
cellRelationship[7:8,] <- "Intimate partner relationships"
cellRelationship[9:10,] <- "Domestic relationships"
cellRelationship[11:12,] <- "Any relationship (known or unknown)"
cellRelationship[13:14,] <- "Intimate partner relationships"
cellRelationship[15:16,] <- "Domestic relationships"
cellRelationship[17:18,] <- "Any relationship (known or unknown)"
cellRelationship[19:20,] <- "Intimate partner relationships"
cellRelationship[21:22,] <- "Domestic relationships"
cellRelationship[23:24,] <- "Any relationship (known or unknown)"
cellRelationship[25:26,] <- "Intimate partner relationships"
cellRelationship[27:28,] <- "Domestic relationships"
cellRelationship[29:30,] <- "Any relationship (known or unknown)"
cellRelationship[31:32,] <- "Intimate partner relationships"
cellRelationship[33:34,] <- "Domestic relationships"
cellRelationship[35:36,] <- "Any relationship (known or unknown)"
cellRelationship[37:38,] <- "Intimate partner relationships"
cellRelationship[39:40,] <- "Domestic relationships"
cellRelationship[41:42,] <- "Any relationship (known or unknown)"
cellRelationship[43:44,] <- "Intimate partner relationships"
cellRelationship[45:46,] <- "Domestic relationships"
cellRelationship[47:48,] <- "Any relationship (known or unknown)"
cellRelationship[49:50,] <- "Intimate partner relationships"
cellRelationship[51:52,] <- "Domestic relationships"
cellRelationship[53:54,] <- "Any relationship (known or unknown)"
cellRelationship[55:58,] <- "Intimate partner relationships"
cellRelationship[59:62,] <- "Domestic relationships"
cellRelationship[63:66,] <- "Any relationship (known or unknown)"
cellRelationship[67:68,] <- "Intimate partner relationships"
cellRelationship[69:70,] <- "Domestic relationships"
cellRelationship[71:72,] <- "Any relationship (known or unknown)"
cellRelationship[73:74,] <- "Intimate partner relationships"
cellRelationship[75:76,] <- "Domestic relationships"
cellRelationship[77:78,] <- "Any relationship (known or unknown)"
cellRelationship[79:80,] <- "Intimate partner relationships"
cellRelationship[81:82,] <- "Domestic relationships"
cellRelationship[83:84,] <- "Any relationship (known or unknown)"

for (i in 1:9)
  cellSex[,i] <- c(rep(c("female victims","total victims"),3),
                 rep(c("offences against female victims","offences against total victims"),3),
                 rep(c("against female victims","against total victims"),3),
                 rep(c("female victims","total victims"),15),
                 rep(c("female victims of femicide","total victims of homicide"),3),
                 rep(c("protection orders with female victims (applied for)","protection orders with total victims (applied for)",
                       "protection orders with female victims (granted)","protection orders with total victims (granted)"),3),
                 rep(c("against female victims","against total victims"),9))

cellYear[,1] <- "2014"
cellYear[,2] <- "2015"
cellYear[,3] <- "2016"
cellYear[,4] <- "2017"
cellYear[,5] <- "2018"
cellYear[,6] <- "2019"
cellYear[,7] <- "2020"
cellYear[,8] <- "2021"
cellYear[,9] <- "2022"

# Read the data from the file
#
# Data are read in blocks of two rows, or, in the case of indicator 10, of four rows.
# However, if such a block is entirely empty, read_excel will not read it.
# To avoid this we always read the row of year above the block, hence the blocks read consist of
# three, or respectively five, rows.
# As soon as each block is read, its top row (the years) is dropped.

adminData <- read_excel(fileName, range = "1__Data!D11:L13",col_names=FALSE)
adminData <- adminData[2:3,]
test1 <- read_excel(fileName, range = "1__Data!D16:L18",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D21:L23",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D29:L31",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D34:L36",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D39:L41",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D47:L49",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D52:L54",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D57:L59",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D65:L67",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D70:L72",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D75:L77",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D83:L85",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D88:L90",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D93:L95",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D101:L103",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D106:L108",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D111:L113",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D119:L121",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D124:L126",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D129:L131",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D137:L139",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D142:L144",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D147:L149",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D155:L157",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D160:L162",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D165:L167",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D176:L180",col_names=FALSE)
test1 <- test1[2:5,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D183:L187",col_names=FALSE)
test1 <- test1[2:5,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D190:L194",col_names=FALSE)
test1 <- test1[2:5,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D200:L202",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D205:L207",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D210:L212",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D218:L220",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D223:L225",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D228:L230",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D236:L238",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D241:L243",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)
test1 <- read_excel(fileName, range = "1__Data!D246:L248",col_names=FALSE)
test1 <- test1[2:3,]
adminData <- rbind(adminData,test1)

rm(test1)

# Create a second tibble, mirroring adminData, but storing the data type of each individual cell
# together with the data. This is useful in operations further down, in the following manner:
# if a cell contains characters all the values in its column are read as characters and the implementation
# of validation rules fails. Having the adminDataTypes tibble we can retrieve each data value with its
# correct data type, irrespectively of the other values' types. Hence, even if there are character values
# in a column, all other values in the column are read as numeric.

adminDataTypes <- read_excel(fileName, range = "1__Data!D11:L13",col_names=FALSE,col_types="list")
adminDataTypes <- adminDataTypes[2:3,]
test1 <- read_excel(fileName, range = "1__Data!D16:L18",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D21:L23",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D29:L31",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D34:L36",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D39:L41",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D47:L49",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D52:L54",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D57:L59",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D65:L67",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D70:L72",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D75:L77",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D83:L85",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D88:L90",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D93:L95",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D101:L103",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D106:L108",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D111:L113",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D119:L121",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D124:L126",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D129:L131",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D137:L139",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D142:L144",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D147:L149",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D155:L157",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D160:L162",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D165:L167",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D176:L180",col_names=FALSE,col_types="list")
test1 <- test1[2:5,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D183:L187",col_names=FALSE,col_types="list")
test1 <- test1[2:5,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D190:L194",col_names=FALSE,col_types="list")
test1 <- test1[2:5,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D200:L202",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D205:L207",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D210:L212",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D218:L220",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D223:L225",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D228:L230",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D236:L238",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D241:L243",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)
test1 <- read_excel(fileName, range = "1__Data!D246:L248",col_names=FALSE,col_types="list")
test1 <- test1[2:3,]
adminDataTypes <- rbind(adminDataTypes,test1)

rm(test1)

# Create an adminDataContents tibble which identifies the cells that are:
# negative (N),
# empty (V),
# character (C)
# real-valued (R)
#
# The unlist(adminDataTypes[i,j][[1]][1] operation returns the value of cell [i,j]
# but having retained the correct data type of it, independently of the types of other
# values in the column. See discussion in the previous block of comments.

adminDataContents <- data.frame(matrix("",nrow=84,ncol=9)) # If blank all is OK with the data
for (i in 1:84)
  for (j in 1:9)
  {
    if (is.na(adminData[i,j])) adminDataContents[i,j] <- "V"
    else if (adminData[i,j]<0) adminDataContents[i,j] <- "N" 
         else if (is.character(unlist(adminDataTypes[i,j][[1]][1]))) adminDataContents[i,j] <- "C"
              else if (unlist(adminDataTypes[i,j][[1]][1]) %% 1 != 0) adminDataContents[i,j] <- "R"
  }
    
# Create the validation report file

WriteFile <- file(valFileName, open = "at", encoding = "native.enc")
text <- paste('VALIDATION REPORT for file: ',fileName,sep="")
writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
text <- paste('NA stands for an empty cell',sep="")
writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)

noErrorsFound <- TRUE

# VALIDATION RULE 1: the cell must contain either a zero or a positive integer, or must be left empty.

if (sum(adminDataContents=="N")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('WARNING! The following cells contain negative numbers:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in 1:84)
    for (j in 1:9)
    {
      if (adminDataContents[i,j]=="N")
        {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
        }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

if (sum(adminDataContents=="R")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('WARNING! The following cells contain positive non-integer numbers:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in 1:84)
    for (j in 1:9)
    {
      if (adminDataContents[i,j]=="R")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

if (sum(adminDataContents=="C")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('ERROR! The following cells contain characters:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in 1:84)
    for (j in 1:9)
    {
      if (adminDataContents[i,j]=="C")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

# VALIDATION RULE 2: the rate of change between consecutive years should not exceed 40% in absolute terms.

# Create temporary data frame for the cells that fail the rule

rule2Failures <- data.frame(matrix("",nrow=84,ncol=9))
for (i in 1:84)
  for (j in 2:9)
  {
    if (adminDataContents[i,j-1]=="" & adminDataContents[i,j]=="") 
    {
      if (adminData[i,j-1]>0 & adminData[i,j]>0)
      {
        t2 <- unlist(adminDataTypes[i,j][[1]][1])
        t1 <- unlist(adminDataTypes[i,j-1][[1]][1])
        if (abs(t2-t1)/t1 > 0.4) rule2Failures[i,j] <- "Warn"
      }
      else if (adminData[i,j-1]>0 | adminData[i,j]>0) rule2Failures[i,j] <- "Warn"
    }
    else
    {
      if (adminDataContents[i,j-1]=="V" & adminDataContents[i,j]=="")
      {
        if (unlist(adminDataTypes[i,j][[1]][1])>0) rule2Failures[i,j] <- "Warn"
      }
      else if (adminDataContents[i,j-1]=="" & adminDataContents[i,j]=="V")
           {
           if (unlist(adminDataTypes[i,j-1][[1]][1])>0) rule2Failures[i,j] <- "Warn"
           }
    }
  }

if (sum(rule2Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('WARNING! In the following pairs of cells the value has changed between consecutive years by more than 40%:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in 1:84)
    for (j in 2:9)
    {
      if (rule2Failures[i,j]=="Warn")
      {
        text <- paste(counter,'. ',cellIndicator[i,j-1],' -- ',cellRelationship[i,j-1],' -- ',cellSex[i,j-1],' -- ',cellYear[i,j-1],' -- ','Value: ',adminData[i,j-1],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

# VALIDATION RULE 3: the number reported for ‘total victims’ should not be lower than the corresponding number reported for ‘female victims’.

# Create temporary data frame for the cells that fail the rule

rule3Failures <- data.frame(matrix("",nrow=84,ncol=9))
for (i in seq(from=1,to=83,by=2))   # These are the rows with data about female victims
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+1,j]=="")
    if ( unlist(adminDataTypes[i+1,j][[1]][1])<unlist(adminDataTypes[i,j][[1]][1])   ) rule3Failures[i,j] <- "Fail"
  }

if (sum(rule3Failures=="Fail")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('ERROR! In the following pairs of cells the value for total victims is lower than the value for female victims:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in seq(from=1,to=83,by=2))
    for (j in 1:9)
    {
      if (rule3Failures[i,j]=="Fail")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+1,j],' -- ',cellRelationship[i+1,j],' -- ',cellSex[i+1,j],' -- ',cellYear[i+1,j],' -- ','Value: ',adminData[i+1,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

# VALIDATION RULE 4: the number reported for ‘domestic relationships’ should not be lower than the corresponding number reported for ‘intimate partner relationships’.

# Create temporary data frame for the cells that fail the rule

rule4Failures <- data.frame(matrix("",nrow=84,ncol=9))
for (i in c(1,2,7,8,13,14,19,20,25,26,31,32,37,38,43,44,49,50,67,68,73,74,79,80))   # These are the rows with data about intimate partner relationships, except for indicator 10
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+2,j]=="")
      if ( unlist(adminDataTypes[i+2,j][[1]][1])<unlist(adminDataTypes[i,j][[1]][1])   ) rule4Failures[i,j] <- "Warn"
  }

for (i in 55:58)   # These are the rows with data about intimate partner relationships for indicator 10
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+4,j]=="")
      if ( unlist(adminDataTypes[i+4,j][[1]][1])<unlist(adminDataTypes[i,j][[1]][1])   ) rule4Failures[i,j] <- "Warn"
  }

if (sum(rule4Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('WARNING! In the following pairs of cells the value for victims in domestic relationships is lower than the value for victims in intimate partner relationships:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in c(1,2,7,8,13,14,19,20,25,26,31,32,37,38,43,44,49,50))
    for (j in 1:9)
    {
      if (rule4Failures[i,j]=="Warn")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+2,j],' -- ',cellRelationship[i+2,j],' -- ',cellSex[i+2,j],' -- ',cellYear[i+2,j],' -- ','Value: ',adminData[i+2,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  for (i in 55:58)
    for (j in 1:9)
    {
      if (rule4Failures[i,j]=="Warn")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+4,j],' -- ',cellRelationship[i+4,j],' -- ',cellSex[i+4,j],' -- ',cellYear[i+4,j],' -- ','Value: ',adminData[i+4,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  for (i in c(67,68,73,74,79,80))
    for (j in 1:9)
    {
      if (rule4Failures[i,j]=="Warn")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+2,j],' -- ',cellRelationship[i+2,j],' -- ',cellSex[i+2,j],' -- ',cellYear[i+2,j],' -- ','Value: ',adminData[i+2,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

# VALIDATION RULE 5: the number reported for ‘any relationship’ should not be lower than the corresponding number reported for ‘domestic relationships’.

# Create temporary data frame for the cells that fail the rule

rule5Failures <- data.frame(matrix("",nrow=84,ncol=9))
for (i in c(3,4,9,10,15,16,21,22,27,28,33,34,39,40,45,46,51,52,69,70,75,76,81,82))   # These are the rows with data about domestic relationships, except for indicator 10
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+2,j]=="")
      if ( unlist(adminDataTypes[i+2,j][[1]][1])<unlist(adminDataTypes[i,j][[1]][1])   ) rule5Failures[i,j] <- "Fail"
  }

for (i in 59:62)   # These are the rows with data about domestic relationships for indicator 10
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+4,j]=="")
      if ( unlist(adminDataTypes[i+4,j][[1]][1])<unlist(adminDataTypes[i,j][[1]][1])   ) rule5Failures[i,j] <- "Fail"
  }

if (sum(rule5Failures=="Fail")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  writeLines('ERROR! In the following pairs of cells the value for victims in any relationship is lower than the value for victims in domestic relationships:',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  for (i in c(3,4,9,10,15,16,21,22,27,28,33,34,39,40,45,46,51,52))
    for (j in 1:9)
    {
      if (rule5Failures[i,j]=="Fail")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+2,j],' -- ',cellRelationship[i+2,j],' -- ',cellSex[i+2,j],' -- ',cellYear[i+2,j],' -- ','Value: ',adminData[i+2,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  for (i in 59:62)
    for (j in 1:9)
    {
      if (rule5Failures[i,j]=="Fail")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+4,j],' -- ',cellRelationship[i+4,j],' -- ',cellSex[i+4,j],' -- ',cellYear[i+4,j],' -- ','Value: ',adminData[i+4,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  for (i in c(69,70,75,76,81,82))
    for (j in 1:9)
    {
      if (rule5Failures[i,j]=="Fail")
      {
        text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        text <- paste(counter,'. ',cellIndicator[i+2,j],' -- ',cellRelationship[i+2,j],' -- ',cellSex[i+2,j],' -- ',cellYear[i+2,j],' -- ','Value: ',adminData[i+2,j],sep="")
        writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
        writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
        counter <- counter+1
      }
    }
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
}

# VALIDATION RULE 6: the number of victims of violence (indicator 1) should not be lower than the corresponding number of victims of physical violence (indicator 4).

# Create temporary data frame for the cells that fail the rule

rule6Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 1:6)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+18,j]=="")
    {
      if ( unlist(adminDataTypes[i+18,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule6Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+18,j]=="") rule6Failures[i,j] <- "Warn"
  }

if (sum(rule6Failures=="Fail")+sum(rule6Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule6Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of victims of any kind of violence (ind. 1) is lower than the corresponding number of victims of physical violence (ind. 4):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule6Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+18,j],' -- ',cellRelationship[i+18,j],' -- ',cellSex[i+18,j],' -- ',cellYear[i+18,j],' -- ','Value: ',adminData[i+18,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule6Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no reported number of victims of any kind of violence (ind. 1) while they have a reported corresponding number of victims of physical violence (ind. 4):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule6Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+18,j],' -- ',cellRelationship[i+18,j],' -- ',cellSex[i+18,j],' -- ',cellYear[i+18,j],' -- ','Value: ',adminData[i+18,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 7: the number of reported offences of violence (indicator 2) should not be lower than the corresponding number of reported offences of physical violence (indicator 4).

# Create temporary data frame for the cells that fail the rule

rule7Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 7:12)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+12,j]=="")
    {
      if ( unlist(adminDataTypes[i+12,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule7Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+12,j]=="") rule7Failures[i,j] <- "Warn"
  }

if (sum(rule7Failures=="Fail")+sum(rule7Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule7Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of reported offences of any kind of violence (ind. 2) is lower than the corresponding number of reported offences of physical violence (ind. 4):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule7Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+12,j],' -- ',cellRelationship[i+12,j],' -- ',cellSex[i+12,j],' -- ',cellYear[i+12,j],' -- ','Value: ',adminData[i+12,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule7Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no data on the number of reported offences of any kind of violence (ind. 2) while they have data on the corresponding number of reported offences of physical violence (ind. 4):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule7Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+12,j],' -- ',cellRelationship[i+12,j],' -- ',cellSex[i+12,j],' -- ',cellYear[i+12,j],' -- ','Value: ',adminData[i+12,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 8: the number of victims of violence (indicator 1) should not be lower than the corresponding number of victims of psychological violence (indicator 5).

# Create temporary data frame for the cells that fail the rule

rule8Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 1:6)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+24,j]=="")
    {
      if ( unlist(adminDataTypes[i+24,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule8Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+24,j]=="") rule8Failures[i,j] <- "Warn"
  }

if (sum(rule8Failures=="Fail")+sum(rule8Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule8Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of victims of any kind of violence (ind. 1) is lower than the corresponding number of victims of psychological violence (ind. 5):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule8Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+24,j],' -- ',cellRelationship[i+24,j],' -- ',cellSex[i+24,j],' -- ',cellYear[i+24,j],' -- ','Value: ',adminData[i+24,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule8Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no reported number of victims of any kind of violence (ind. 1) while they have a reported corresponding number of victims of psychological violence (ind. 5):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule8Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+24,j],' -- ',cellRelationship[i+24,j],' -- ',cellSex[i+24,j],' -- ',cellYear[i+24,j],' -- ','Value: ',adminData[i+24,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 9: the number of reported offences of violence (indicator 2) should not be lower than the corresponding number of reported offences of psychological violence (indicator 5).

# Create temporary data frame for the cells that fail the rule

rule9Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 7:12)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+18,j]=="")
    {
      if ( unlist(adminDataTypes[i+18,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule9Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+18,j]=="") rule9Failures[i,j] <- "Warn"
  }

if (sum(rule9Failures=="Fail")+sum(rule9Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule9Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of reported offences of any kind of violence (ind. 2) is lower than the corresponding number of reported offences of psychological violence (ind. 5):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule9Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+18,j],' -- ',cellRelationship[i+18,j],' -- ',cellSex[i+18,j],' -- ',cellYear[i+18,j],' -- ','Value: ',adminData[i+18,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule9Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no data on the number of reported offences of any kind of violence (ind. 2) while they have data on the corresponding number of reported offences of psychological violence (ind. 5):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule9Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+18,j],' -- ',cellRelationship[i+18,j],' -- ',cellSex[i+18,j],' -- ',cellYear[i+18,j],' -- ','Value: ',adminData[i+18,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 10: the number of victims of violence (indicator 1) should not be lower than the corresponding number of victims of sexual violence (indicator 6).

# Create temporary data frame for the cells that fail the rule

rule10Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 1:6)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+30,j]=="")
    {
      if ( unlist(adminDataTypes[i+30,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule10Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+30,j]=="") rule10Failures[i,j] <- "Warn"
  }

if (sum(rule10Failures=="Fail")+sum(rule10Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule10Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of victims of any kind of violence (ind. 1) is lower than the corresponding number of victims of sexual violence (ind. 6):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule10Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+30,j],' -- ',cellRelationship[i+30,j],' -- ',cellSex[i+30,j],' -- ',cellYear[i+30,j],' -- ','Value: ',adminData[i+30,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule10Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no reported number of victims of any kind of violence (ind. 1) while they have a reported corresponding number of victims of sexual violence (ind. 6):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule10Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+30,j],' -- ',cellRelationship[i+30,j],' -- ',cellSex[i+30,j],' -- ',cellYear[i+30,j],' -- ','Value: ',adminData[i+30,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 11: the number of reported offences of violence (indicator 2) should not be lower than the corresponding number of reported offences of sexual violence (indicator 6).

# Create temporary data frame for the cells that fail the rule

rule11Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 7:12)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+24,j]=="")
    {
      if ( unlist(adminDataTypes[i+24,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule11Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+24,j]=="") rule11Failures[i,j] <- "Warn"
  }

if (sum(rule11Failures=="Fail")+sum(rule11Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule11Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of reported offences of any kind of violence (ind. 2) is lower than the corresponding number of reported offences of sexual violence (ind. 6):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule11Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+24,j],' -- ',cellRelationship[i+24,j],' -- ',cellSex[i+24,j],' -- ',cellYear[i+24,j],' -- ','Value: ',adminData[i+24,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule11Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no data on the number of reported offences of any kind of violence (ind. 2) while they have data on the corresponding number of reported offences of sexual violence (ind. 6):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule11Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+24,j],' -- ',cellRelationship[i+24,j],' -- ',cellSex[i+24,j],' -- ',cellYear[i+24,j],' -- ','Value: ',adminData[i+24,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 12: the number of victims of violence (indicator 1) should not be lower than the corresponding number of victims of economic violence (indicator 7).

# Create temporary data frame for the cells that fail the rule

rule12Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 1:6)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+36,j]=="")
    {
      if ( unlist(adminDataTypes[i+36,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule12Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+36,j]=="") rule12Failures[i,j] <- "Warn"
  }

if (sum(rule12Failures=="Fail")+sum(rule12Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule12Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of victims of any kind of violence (ind. 1) is lower than the corresponding number of victims of economic violence (ind. 7):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule12Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+36,j],' -- ',cellRelationship[i+36,j],' -- ',cellSex[i+36,j],' -- ',cellYear[i+36,j],' -- ','Value: ',adminData[i+36,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule12Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no reported number of victims of any kind of violence (ind. 1) while they have a reported corresponding number of victims of economic violence (ind. 7):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 1:6)
      for (j in 1:9)
      {
        if (rule12Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+36,j],' -- ',cellRelationship[i+36,j],' -- ',cellSex[i+36,j],' -- ',cellYear[i+36,j],' -- ','Value: ',adminData[i+36,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 13: the number of reported offences of violence (indicator 2) should not be lower than the corresponding number of reported offences of economic violence (indicator 7).

# Create temporary data frame for the cells that fail the rule

rule13Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 7:12)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+30,j]=="")
    {
      if ( unlist(adminDataTypes[i+30,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule13Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+30,j]=="") rule13Failures[i,j] <- "Warn"
  }

if (sum(rule13Failures=="Fail")+sum(rule13Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule13Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the number of reported offences of any kind of violence (ind. 2) is lower than the corresponding number of reported offences of economic violence (ind. 7):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule13Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+30,j],' -- ',cellRelationship[i+30,j],' -- ',cellSex[i+30,j],' -- ',cellYear[i+30,j],' -- ','Value: ',adminData[i+30,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule13Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no data on the number of reported offences of any kind of violence (ind. 2) while they have data on the corresponding number of reported offences of economic violence (ind. 7):',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 7:12)
      for (j in 1:9)
      {
        if (rule13Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+30,j],' -- ',cellRelationship[i+30,j],' -- ',cellSex[i+30,j],' -- ',cellYear[i+30,j],' -- ','Value: ',adminData[i+30,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# VALIDATION RULE 14: the number of victims of sexual violence (indicator 6) should not be lower than the corresponding number of victims reporting rape (indicator 8).
# For Member States that report number of offences in indicators 6 and 8 the rule stays the same, but refers to number of offences.

# Create temporary data frame for the cells that fail the rule

rule14Failures <- data.frame(matrix("",nrow=84,ncol=9))

for (i in 31:36)
  for (j in 1:9)
  {
    if (adminDataContents[i,j]=="" & adminDataContents[i+12,j]=="")
    {
      if ( unlist(adminDataTypes[i+12,j][[1]][1])>unlist(adminDataTypes[i,j][[1]][1])   ) rule14Failures[i,j] <- "Fail"
    }
    else if (adminDataContents[i,j]=="V" & adminDataContents[i+12,j]=="") rule14Failures[i,j] <- "Warn"
  }

if (sum(rule14Failures=="Fail")+sum(rule14Failures=="Warn")>0)
{
  noErrorsFound <- FALSE
  counter <- 1
  if (sum(rule14Failures=="Fail")>0)
  {
    writeLines('ERROR! In the following pairs of cells the value of indicator 6 is lower than the corresponding value of indicator 8:',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 31:36)
      for (j in 1:9)
      {
        if (rule14Failures[i,j]=="Fail")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+12,j],' -- ',cellRelationship[i+12,j],' -- ',cellSex[i+12,j],' -- ',cellYear[i+12,j],' -- ','Value: ',adminData[i+12,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
  if (sum(rule14Failures=="Warn")>0)
  {
    writeLines('WARNING! The following pairs of cells have no data for indicator 6 while they have the corresponding data for indicator 8:',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    for (i in 31:36)
      for (j in 1:9)
      {
        if (rule14Failures[i,j]=="Warn")
        {
          text <- paste(counter,'. ',cellIndicator[i,j],' -- ',cellRelationship[i,j],' -- ',cellSex[i,j],' -- ',cellYear[i,j],' -- ','Value: ',adminData[i,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          text <- paste(counter,'. ',cellIndicator[i+12,j],' -- ',cellRelationship[i+12,j],' -- ',cellSex[i+12,j],' -- ',cellYear[i+12,j],' -- ','Value: ',adminData[i+12,j],sep="")
          writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
          writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
          counter <- counter+1
        }
      }
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
    writeLines('   ',sep="\n",con = WriteFile,useBytes = TRUE)
  }
}

# Still write something in the validation output when nothing has been detected by validation

if (noErrorsFound==TRUE)
{
  text<-paste('No errors or warnings were generated by the automated validation of the file!')
  writeLines(text,sep="\n",con = WriteFile,useBytes = TRUE)
}

# Close the validation report file

close(WriteFile)
