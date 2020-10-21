
library(readxl) 
library(XLConnect)
data_list <- import_list("C:/Users/ADOOR PRABHA/Desktop/recent_sheets.xlsx", setclass = "tbl")
t<-excel_sheets("C:/Users/ADOOR PRABHA/Desktop/recent_sheets.xlsx")
wb1<-loadWorkbook("theoverlap6.xlsx",create = TRUE)
for(i in 1:length(t))
{
  createSheet(wb1,toString(t[i]))
  #y<-colnames(data_list[[i]])
  #print(y)
  for(j in 2:ncol(data_list[[i]]))
  {
    v<-data_list[[i]][[1]]
    q<-data_list[[i]][[j]]
    w<-data.frame(intersect(v,q))
    writeWorksheet(wb1,w,toString(t[i]),startRow =1 ,startCol =j ,header = TRUE,rownames = TRUE)
    
    
  }
 saveWorkbook(wb1,"theoverlap6.xlsx")
}
