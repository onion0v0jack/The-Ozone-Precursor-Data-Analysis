rm(list=ls())  #清除所有暫存
graphics.off() #清除圖片的顯示
getwd()        #顯示這個r所在的路徑

Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_144') 
require(ggplot2)
library(ggfortify)
library(readxl)
require(plyr)
require(dplyr)
require(RColorBrewer)
require(NMF)
setwd("C:/Users/user/下載/中山應數/Meeting/論文/光化/106年前驅物資料/日平均資料")
list_file=list.files(pattern ="*.xlsx")

list_A=list()
for(i in 1:length(list_file))
{list_A[[i]]=readxl::read_excel(list_file[[i]],col_names=TRUE)}

which(is.na(list_A[[9]]),arr.ind =T) #看是否有NA
  

###前驅物權重###
r=1
set.seed(1)
ImportantPLOT=list()
for(j in c(1,2,3,4,7,8,9)) #1:length(list_file)
{
  Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
  Height=c(200,400,600,776); CexNames=c(0.95,0.95,1.4,1.4); CexMain=c(2,2,2.75,3); CexLab=c(1.3,1.3,1.8,1.8)
  res <- nmf(Data, rank=r)
  png(filename=paste(substr(list_file[j],start=1,stop=5),".png",sep=""), width=1365, height=Height[r])
  par(mfrow=c(r,1))
  par(mar = c(5.1,5,4.1,0)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  for (base in 1:r)
  {
    temp=res@fit@W[,base]
    names(temp)=1:54
    barplot(temp,cex.names=CexNames[r],cex.lab=CexLab[r],cex.main=CexMain[r],
            ylab="Loading",xlab="Number of precursor",main=paste("Basis",base))
  }
  dev.off();#
}

###前驅物權重 按照大小排序###
#順便選取前驅物(以排序後最大差作為門檻)
r=4
set.seed(1)
ImportantPLOT1=list()
for(j in 1:length(list_file)) #1:length(list_file)
{
  Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
  Height=c(200,400,600,776); CexNames=c(0.95,0.95,1.4,1.4); CexMain=c(2,2,2.75,3); CexLab=c(1.3,1.3,1.8,1.8)
  res <- nmf(Data, rank=r)
  png(filename=paste(substr(list_file[j],start=1,stop=5),".png",sep=""), width=1365, height=Height[r])
  par(mfrow=c(r,1))
  par(mar = c(5.1,5,4.1,0)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  for (base in 1:r)
  {
    temp=res@fit@W[rev(order(res@fit@W[,base])),base]
    names(temp)=rev(order(res@fit@W[,base]))
    barplot(temp,cex.names=CexNames[r],cex.lab=CexLab[r],cex.main=CexMain[r],
            ylab="Loading",xlab="Number of precursor",main=paste("Basis",base))
  }
  dev.off();#
  PLOT1=NULL
  for(i in 1:r)
  {
    OrderPLOT=rev(order(res@fit@W[,i]))#排名由大到小
    Order=res@fit@W[OrderPLOT,i]
    PLOT1=c(PLOT1,OrderPLOT[1:order(diff(array(Order)))[1]])
  }
  PLOT1=unique(sort(PLOT1))
  ImportantPLOT1[[j]]=PLOT1
  print(paste("St",j))
  print(PLOT1)
  
  
}

######
#setwd("C:/Users/user/下載/中山應數/Meeting/論文/光化")
#PLOTsource=read.csv("臭氧前驅物排放源整理表.csv")
#st=1
#PLOTsource=cbind(rep(0,54),PLOTsource); names(PLOTsource)[1]="Important_PLOT"
#PLOTsource[ImportantPLOT[[st]],1]=1
#PLOTsource%>%filter(Important_PLOT==1)%>%View()
#View(PLOTsource)

