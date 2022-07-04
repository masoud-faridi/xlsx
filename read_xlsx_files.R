library(openxlsx)

filenamerv<-file.choose()
sheetsrv <- getSheetNames(filenamerv)
SheetListrv <- lapply(sheetsrv,read.xlsx,xlsxFile=filenamerv)
names(SheetListrv) <- sheetsrv
names(SheetListrv)


filename97<-file.choose()
sheets97 <- openxlsx::getSheetNames(filename97)
SheetList97 <- lapply(sheets97,openxlsx::read.xlsx,xlsxFile=filename97)
names(SheetList97) <- sheets97
names(SheetList97)

filename98<-file.choose()
sheets98 <- openxlsx::getSheetNames(filename98)
SheetList98 <- lapply(sheets98,openxlsx::read.xlsx,xlsxFile=filename98)
names(SheetList98) <- sheets98
n<-length(names(SheetList98))

temp<-rbind(as.data.frame(SheetList97[1])[-13,],as.data.frame(SheetList98[1])[-13,])

exportexcel<-rbind(as.data.frame(SheetList97[1])[-13,],as.data.frame(SheetList98[1])[-13,])

for(i in 2:n){
temp<-rbind(as.data.frame(SheetList97[i])[-13,],as.data.frame(SheetList98[i])[-13,])
exportexcel<-rbind(exportexcel,temp)
}



temp<-rbind(as.matrix(SheetList97[1])[-13,],as.matrix(SheetList98[1])[-13,])
exportexcel<-rbind(as.matrix(SheetList97[1])[-13,],as.matrix(SheetList98[1])[-13,])
for(i in 2:n){
temp<-rbind(as.matrix(SheetList97[i])[-13,],as.matrix(SheetList98[i])[-13,])
exportexcel<-rbind(exportexcel,temp)
}
write.xlsx(exportexcel,"dataexport0.xlsx")



do.call("rbind", list(DF1, DF2, DF3))



temp<-rbind(rbind(as.data.frame(SheetList97[1])[-13,],as.data.frame(SheetList98[1])[-13,]))

exportexcel<-list()
exportexcel[[1]]<-temp
for(i in 2:n){
temp<-rbind(as.data.frame(SheetList97[i])[-13,],as.data.frame(SheetList98[i])[-13,])
exportexcel[[i]]<-temp
#print(dim(temp))
#readline()
}
write.xlsx(exportexcel,"dataexport10.xlsx")


mat<-matrix(0,ncol=
for(i in 1:n){
temp<-rbind(as.data.frame(SheetList97[i])[-13,],as.data.frame(SheetList98[i])[-13,])
dim(temp)
}





library(openxlsx)
path <- "path/to/your/file.xlsx"
getSheetNames(path)




library(openxlsx)
wb <- createWorkbook()  #wb <- loadWorkbook("RawExcel.xlsx")
addWorksheet(wb, sheetName = "sheetname1")
writeData(wb, sheet = "sheetname1", x = data_sheetname1)
addWorksheet(wb, sheetName = "sheetname2")
writeData(wb, sheet = "sheetname2", x = data_sheetname2)
saveWorkbook(wb, "D:\\speed.xlsx")

file_choose<-file.choose()
data_output<-read.csv(file_choose,encoding = "UTF-8")
View(data_output)
data_sheetname1<-data_output










convert_csv_to_xlsx<-function(){
  #options(encoding = "UTF-8")
  filename97<-file.choose()
  datacc<-read.csv(filename97,sep=",",encoding="UTF-8")
  View(datacc)
  library(openxlsx)
  wb <- createWorkbook() 
  addWorksheet(wb, sheetName = "sheetname1")
  writeData(wb, sheet = "sheetname1", x = datacc)
  saveWorkbook(wb, "converted.xlsx")
  
}








data_clean<-function(sheetname="tasadof"){
 # yefarsi<-"ی"
  yefarsi<-"\U06CC"
  #as.character(yefarsi)
  #cat(yefarsi)
  #yearabic<-"ي"
  yearabic<-"\u064A"
  #kehfarsi<-"ک"
  kehfarsi<-"\u06A9"
  #keyarabi<-"ك"
  keyarabi<-"\u0643"
  #  gsub(yearabic,yefarsi,"گيلان")
  library(openxlsx)
  setwd(choose.dir())
  filename_tasadofat<-file.choose()
  # P_data_tasadofat <- openxlsx::read.xlsx(filename_tasadofat, sheet= "tasadof")
  P_data_tasadofat <- openxlsx::read.xlsx(filename_tasadofat, sheet= sheetname)
  P_data_tasadofat<-as.data.frame(P_data_tasadofat,encoding="UTF-8")
 
  P_data_tasadofat$OSTAN<-gsub(yearabic,yefarsi,P_data_tasadofat$OSTAN)
  P_data_tasadofat$OSTAN<-gsub(keyarabi,kehfarsi,P_data_tasadofat$OSTAN)
  
  
# if(sheetname=="tasadof"){
  # k_va_b<-"كهکیلویه و بویراحمد"
  k_va_b<-"\u06A9\u0647\u06A9\u06CC\u0644\u0648\u06CC\u0647\u20\u0648\u20\u0628\u0648\u06CC\u0631\u0627\u062D\u0645\u062F"
 # %u06A9%u0647%u06A9%u06CC%u0644%u0648%u06CC%u0647%20%u0648%20%u0628%u0648%u06CC%u0631%u0627%u062D%u0645%u062F
  #correct_k_va_b<-"کهگیلویه و بویراحمد"
  correct_k_va_b<-"\u06A9\u0647\u06AF\u06CC\u0644\u0648\u06CC\u0647\u20\u0648\u20\u0628\u0648\u06CC\u0631\u0627\u062D\u0645\u062F"
 # which(P_data_tasadofat$OSTAN== k_va_b)
 # which(P_data_tasadofat$OSTAN=="کهکیلویه و بویراحمد")
 # stringdist("کهکیلویه و بویراحمد",k_va_b)
  
  P_data_tasadofat$OSTAN<-gsub( k_va_b,correct_k_va_b,P_data_tasadofat$OSTAN)
#   }
#  if(sheetname=="taradod"){
  #chob<-"چهارمحال و بختیاری"
  chob<-"\u0686\u0647\u0627\u0631\u0645\u062D\u0627\u0644\u20\u0648\u20\u0628\u062E\u062A\u06CC\u0627\u0631\u06CC"
  # correct_chob<-"چهار محال و بختیاری"
  correct_chob<-"\u0686\u0647\u0627\u0631\u20\u0645\u062D\u0627\u0644\u20\u0648\u20\u0628\u062E\u062A\u06CC\u0627\u0631\u06CC"
  P_data_tasadofat$OSTAN<-gsub( chob,correct_chob,P_data_tasadofat$OSTAN)
#   }
 
  #View(P_data_tasadofat)
  exportdata<-list(P_data_tasadofat)
  names(exportdata)<-sheetname
    write.xlsx(exportdata,paste(sheetname,"clean.xlsx"))
 
}