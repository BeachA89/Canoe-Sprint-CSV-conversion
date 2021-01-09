# Load libraries
library(tidyverse)
library(svDialogs)
library(xlsx)
library(dplyr)


remove(list = ls())
#Import file with dialog window

TxtData <-  dlg_open(default= "C:/Users/aaron.beach/Dropbox/Kayak 2019-2020/NSWIS 2020 competitions/Nationals/Raw Excels", Title = "Select a  file")$res

#Read the xlsx sheet one that contains race information

data1 <-  read.xlsx(TxtData, 1, stringsAsFactors = FALSE, header = FALSE)

#Read the xlsx sheet 9 that contact environmental data for all races

Environmentals <-  read.xlsx(TxtData, 9, stringsAsFactors = FALSE, header = FALSE)

Temp1 = Environmentals[5,3]
Temp2 = Environmentals[18,3]
Temp3 = Environmentals[30,3]

Humidity1 = Environmentals[5,7]
Humidity2 = Environmentals[18,7]
Humidity3 = Environmentals[30,7]

WindSpeed1 = Environmentals[8,3]
WindSpeed2 = Environmentals[21,3]
WindSpeed3 = Environmentals[33,3]

WindDir1 = Environmentals[8,7]
WindDir2 = Environmentals[21,7]
WindDir3 = Environmentals[33,7]

#Convert labels to Pipeline labels

if (WindDir1 ==  'Cross Head - L') {
  WindDir1 =  "Cross_head"
} else if (WindDir1 == 'Cross Head - R') {
  WindDir1 = "Cross_head"
} else if (WindDir1 ==  'Cross Tail - L') {
  WindDir1 =  "Cross_tail"
} else if (WindDir1 == 'Cross Tail - R'){
  WindDir1 =  "Cross_tail"
}

if (WindDir2 ==  'Cross Head - L') {
  WindDir2 =  "Cross_head"
} else if (WindDir2 == 'Cross Head - R') {
  WindDir2 = "Cross_head"
} else if (WindDir2 ==  'Cross Tail - L') {
  WindDir2 =  "Cross_tail"
} else if (WindDir2 == 'Cross Tail - R'){
  WindDir2 =  "Cross_tail"
}

if (WindDir3 ==  'Cross Head - L') {
  WindDir3 =  "Cross_head"
} else if (WindDir3 == 'Cross Head - R') {
  WindDir3 = "Cross_head"
} else if (WindDir3 ==  'Cross Tail - L') {
  WindDir3 =  "Cross_tail"
} else if (WindDir3 == 'Cross Tail - R'){
  WindDir3 =  "Cross_tail"
}




#Add names with dialog window (will have this as an auto step soon)
First_Last <- dlgInput("Enter First_Last Name for all athletes")$res


# Name <-  data1[2,4]
# Namesplit <- str_split(Name, " ")

#Identify metadata labels, some entered here since they are common/consistent

Event <- "NationalChampionships"
Location <-  "SIRC"
Date <- "11032020"
# convert Gender to Pipeline labels (M/W)
Gender <-  data1[11,4]
if (Gender ==  'Men') {
  Gender =  "M"
} else if (Gender == 'Women') {
  Gender = "W"
}


Distance <-  data1[9,4]
Class <-  data1[10,4]
Age <-  data1[12,4]
conversionfactor <-  0.0000115740740740741

#Read the xlsx sheet two that contains 1st race data

data2 <-  read.xlsx(TxtData, 2, stringsAsFactors = FALSE, header = FALSE)

#Read the xlsx sheet three that contains 2nd race data

data3 <-  read.xlsx(TxtData, 3, stringsAsFactors = FALSE, header = FALSE)

#Read the xlsx sheet four that contains 3rd race data

data4 <-  read.xlsx(TxtData, 4, stringsAsFactors = FALSE, header = FALSE)

#####################################

# Process Race 1


Race <-  data2[17,4]
if (Race ==  'Heat') {
  Race =  "H"
} else if (Race == 'Semi-Final') {
  Race = "SF"
} else if (Race ==  'A Final') {
  Race =  "FA"
} else if (Race == 'B Final'){
  Race =  "FB"
} else if (Race ==  'C Final') {
  Race[1] = "FC"
}

#Identify metadata
Lane <-  data2[18,4]
Place <-  data2[19,4]
OfficialTime <-  (as.numeric(data2[20,4]))/conversionfactor
WinningTime <-  (as.numeric(data2[21,4]))/conversionfactor

#extract racedata to dataframe
racedata <- data2[35:nrow(data2),1:6]

# convert race time
if (as.numeric(racedata[1,6]) < 1) 
{racedata[1,6] <- (as.numeric(racedata[1,6]))/conversionfactor
} else {
  racedata[1,6] <- racedata[1,6]
}

#extract interval time to convert to numeric and then loop to convert time about 1min
intervaltime <- racedata[3:nrow(racedata),2]
intervaltime <- as.numeric(intervaltime)
intervaltime <-  na.omit(intervaltime)


for (i in 1:length(intervaltime)) {
  if (intervaltime[i] < 1)
  {intervaltime[i]= intervaltime[i]/conversionfactor
  }   else {
    intervaltime[i] =  intervaltime[i]
  }
}

# import back to racedataframe
racedata[3:(length(intervaltime)+2),2] <-  intervaltime



#create specified filename i.e. CSVname = First_Last_GrandPrix2_DDMMYYY_SIRC_MK1_500_Senior_FinalA.csv
directory = "C:/Users/aaron.beach/OneDrive - nswis.com.au/Projects/Pipeline project/Phase 2 - Canoe Sprint/CSV Dump R Nationals/"
Racename = paste0(First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race)
file <- paste0(directory,First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race,".csv")

write.table(racedata, file, na = "", sep=",", row.names = FALSE, col.names=FALSE)

Athleterace = paste0(directory,"meta_",First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,".csv")

# Write metadata to a file

a <-  data.frame(Racename, Lane, Place, OfficialTime, WinningTime, Temp1, Humidity1, WindSpeed1, WindDir1,"", "",OfficialTime,WinningTime)
metadata <- rbind(a)
write.table(metadata, Athleterace, sep=",", na="", row.names = FALSE, col.names=FALSE)

##############################################

#Process Race 2

Race <-  data3[17,4]
if (Race ==  'Heat') {
  Race =  "H"
} else if (Race == 'Semi-Final') {
  Race = "SF"
} else if (Race ==  'A Final') {
  Race =  "FA"
} else if (Race == 'B Final'){
  Race =  "FB"
} else if (Race ==  'C Final') {
  Race[1] = "FC"
}


Lane <-  data3[18,4]
Place <-  data3[19,4]
OfficialTime <-  (as.numeric(data3[20,4]))/conversionfactor
WinningTime <-  (as.numeric(data3[21,4]))/conversionfactor

#extract racedata to dataframe
racedata <- data3[35:nrow(data3),1:6]

# convert race time
if (as.numeric(racedata[1,6]) < 1) 
{racedata[1,6] <- (as.numeric(racedata[1,6]))/conversionfactor
} else {
  racedata[1,6] <- racedata[1,6]
}

#extract interval time to convert to numeric and then loop to convert time about 1min
intervaltime <- racedata[3:nrow(racedata),2]
intervaltime <- as.numeric(intervaltime)
intervaltime <-  na.omit(intervaltime)


for (i in 1:length(intervaltime)) {
  if (intervaltime[i] < 1)
  {intervaltime[i]= intervaltime[i]/conversionfactor
  }   else {
    intervaltime[i] =  intervaltime[i]
  }
}

# import back to racedataframe
racedata[3:(length(intervaltime)+2),2] <-  intervaltime


#create specified filename i.e. CSVname = First_Last_GrandPrix2_DDMMYYY_SIRC_MK1_500_Senior_FinalA.csv
file <- paste0(directory,First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race,".csv")
Racename = paste0(First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race)


write.table(racedata, file, na = "", sep=",", row.names = FALSE, col.names=FALSE)

# Write metadata to a file (verwrite previous if there is a Race 2)
b <-  data.frame(Racename, Lane, Place, OfficialTime, WinningTime, Temp2, Humidity2, WindSpeed2, WindDir2,"", "",OfficialTime,WinningTime)
names(b)<-names(a)

metadata <- rbind(a,b)
write.table(metadata, Athleterace, na="", sep=",", row.names = FALSE, col.names=FALSE)

###################################################################

# Process Race 3

Race <-  data4[17,4]
if (Race ==  'Heat') {
  Race =  "H"
} else if (Race == 'Semi-Final') {
  Race = "SF"
} else if (Race ==  'A Final') {
  Race =  "FA"
} else if (Race == 'B Final'){
  Race =  "FB"
} else if (Race ==  'C Final') {
  Race[1] = "FC"
}


Lane <-  data4[18,4]
Place <-  data4[19,4]
OfficialTime <-  (as.numeric(data4[20,4]))/conversionfactor
WinningTime <-  (as.numeric(data4[21,4]))/conversionfactor

#extract racedata to dataframe
racedata <- data4[35:nrow(data4),1:6]

# convert race time
if (as.numeric(racedata[1,6]) < 1) 
{racedata[1,6] <- (as.numeric(racedata[1,6]))/conversionfactor
} else {
  racedata[1,6] <- racedata[1,6]
}

#extract interval time to convert to numeric and then loop to convert time about 1min
intervaltime <- racedata[3:nrow(racedata),2]
intervaltime <- as.numeric(intervaltime)
intervaltime <-  na.omit(intervaltime)


for (i in 1:length(intervaltime)) {
  if (intervaltime[i] < 1)
  {intervaltime[i]= intervaltime[i]/conversionfactor
  }   else {
    intervaltime[i] =  intervaltime[i]
  }
}

# import back to racedataframe
racedata[3:(length(intervaltime)+2),2] <-  intervaltime


###################

#create specified filename i.e. CSVname = First_Last_GrandPrix2_DDMMYYY_SIRC_MK1_500_Senior_FinalA.csv

file <- paste0(directory,First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race,".csv")
Racename = paste0(First_Last,"_",Event,"_",Date,"_",Location,"_",Gender,Class,"_",Distance,"_",Age,"_", Race)


write.table(racedata, file, na = "", sep=",", row.names = FALSE, col.names=FALSE)

c <-  data.frame(Racename, Lane, Place, OfficialTime, WinningTime, Temp3, Humidity3, WindSpeed3, WindDir3,"", "",OfficialTime,WinningTime)

names(c)<-names(a)

# Write metadata to a file (overwrite previous if there is a Race 3)

metadata <- rbind(a,b,c)
write.csv(metadata, Athleterace, na = "", row.names = FALSE, col.names=FALSE)