rm(list=ls())  #�M���Ҧ��Ȧs
graphics.off() #�M���Ϥ������
getwd()        #��ܳo��r�Ҧb�����|

Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_144') 
require(ggplot2)
library(ggfortify)
library(readxl)
require(plyr)
require(dplyr)
require(RColorBrewer)
require(NMF)
setwd("C:/Users/user/�U��/���s����/Meeting/�פ�/����/106�~�e�X�����/�饭�����")
list_file=list.files(pattern ="*.xlsx")

list_A=list()
for(i in 1:length(list_file))
{list_A[[i]]=readxl::read_excel(list_file[[i]],col_names=TRUE)}

which(is.na(list_A[[9]]),arr.ind =T) #�ݬO�_��NA
  

###�e�X���v��###
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
  par(mar = c(5.1,5,4.1,0)) #�]�w�ϧΪť���ɦ�ƤU���W�k�ADefault=c(5.1, 4.1, 4.1, 2.1)
  for (base in 1:r)
  {
    temp=res@fit@W[,base]
    names(temp)=1:54
    barplot(temp,cex.names=CexNames[r],cex.lab=CexLab[r],cex.main=CexMain[r],
            ylab="Loading",xlab="Number of precursor",main=paste("Basis",base))
  }
  dev.off();#
}

###�e�X���v�� ���Ӥj�p�Ƨ�###
#���K����e�X��(�H�Ƨǫ�̤j�t�@�����e)
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
  par(mar = c(5.1,5,4.1,0)) #�]�w�ϧΪť���ɦ�ƤU���W�k�ADefault=c(5.1, 4.1, 4.1, 2.1)
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
    OrderPLOT=rev(order(res@fit@W[,i]))#�ƦW�Ѥj��p
    Order=res@fit@W[OrderPLOT,i]
    PLOT1=c(PLOT1,OrderPLOT[1:order(diff(array(Order)))[1]])
  }
  PLOT1=unique(sort(PLOT1))
  ImportantPLOT1[[j]]=PLOT1
  print(paste("St",j))
  print(PLOT1)
  
  
}

######
#setwd("C:/Users/user/�U��/���s����/Meeting/�פ�/����")
#PLOTsource=read.csv("���e�X���Ʃ񷽾�z��.csv")
#st=1
#PLOTsource=cbind(rep(0,54),PLOTsource); names(PLOTsource)[1]="Important_PLOT"
#PLOTsource[ImportantPLOT[[st]],1]=1
#PLOTsource%>%filter(Important_PLOT==1)%>%View()
#View(PLOTsource)
