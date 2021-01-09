library(dplyr)
library(readr)
library(data.table)

mypath <- "C:/Users/aaron.beach/OneDrive - nswis.com.au/Projects/Pipeline project/Phase 2 - Canoe Sprint/CSV Dump R State Champs/meta"



multmerge = function(path){
  filenames=list.files(path=path, full.names=TRUE)
  rbindlist(lapply(filenames, fread))
}


DF <- multmerge(mypath)
Combined = paste0(mypath,"/combinedmeta",".csv")
write.table(DF, Combined, sep=",", row.names = FALSE, col.names=FALSE)
