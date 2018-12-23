rm(list=ls())  #清除所有暫存
graphics.off() #清除圖片的顯示
getwd()        #顯示這個r所在的路徑

Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_144') 
require(ggplot2)
library(ggfortify)
library(readxl)
require(RColorBrewer)
require(NMF)
setwd("C:/Users/user/下載/中山應數/Meeting/論文/光化/106年前驅物資料/日平均資料")
list_file=list.files(pattern ="*.xlsx")#[-9]; #傳回指定路徑中的檔案名稱字元向量或子目錄。

list_A=list(); #把檔名在list_file中的檔案蒐集成一個list
for(i in 1:length(list_file))
{
  list_A[[i]]=readxl::read_excel(list_file[[i]],col_names = TRUE)
}

which(is.na(list_A[[9]]),arr.ind =T) #看是否有NA

###########以下不用跑#################
mod_A=list()
for(j in 1:length(list_file))
{
  j=1
  A=list_A[[j]]
  mod_A[[j]]=A;
  for(i in 2:ncol(A))#因為每一個資料的第一個col都代表日期
  {
    marfit=arima(A[,i],order = c(1,0,0) ); #A[,i]表示每一個資料的臭氧前驅物(最後一個為溫度)
    mod_A[[j]][,i]=marfit$residuals; #利用AR(1)配適每一個前驅物的日平均數值，並且將其殘差存在mod_A中
  }
}#mod_A中所有資料的size不變，當中的日期也沒有變，只有其他數值殘差而已

#Loading(使用到殘差，也就是mod_A的資料)
for(j in 1:length(list_file))
{
  A1=mod_A[[j]][,-1*c(1,ncol(mod_A[[j]]))];#扣掉資料中的日期跟溫度不計
  pca <- prcomp( A1, scale = FALSE) 
  
  png(filename = paste(list_file[j],"loading",".png",sep=""),
      width = 2048, height = 1024) #建立一個png並設定檔名，但圖的內容還需要下面決定
  par(mfrow=c(2,1))
  tmp=pca$rotation[,1]; #第一主成分(此向量對於各前驅物的權重)
  names(tmp)=1:length(tmp);
  barplot(tmp,cex.names=1.25,xlab="PCA1 Loading",main=paste(gsub(".xlsx","",list_file[j]),"De-Time Series"))
  
  tmp=pca$rotation[,2]; #第二主成分
  names(tmp)=1:length(tmp);
  barplot(tmp,cex.names=1.25,xlab="PCA2 Loading")
  dev.off();
  
  #把第一主成分跟第二主成分建立一個平面，並把各前驅物對應其在兩主成分的分量，顯示出來
  autoplot(pca, data = A1,
           loadings = TRUE, loadings.colour = 'blue',
           loadings.label = TRUE, loadings.label.size = 4,main=paste(gsub(".xlsx","",list_file[j]),"De-Time Series"))
  ggsave( paste(list_file[j],"loading_scores",".png",sep="")) #存檔
}
###########以上不用跑#################

#original(使用原本資料 也就是list_A)
PCA1weight=PCA2weight=as.data.frame(matrix(NA,54,9))
colnames(PCA1weight)=colnames(PCA2weight)=c("S1","S2","S3","S4","S5","S6","S7","S8","S9")
for(j in 1:length(list_file))
{
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #去掉日期與溫度
  pca <- prcomp( A1, scale = FALSE) #建立pca
  
  #png(filename = paste(list_file[j],"loading_original",".png",sep=""),
  #    width = 2048, height = 1024)
  #par(mfrow=c(2,1))
  tmp=pca$rotation[,1]; #第一主成分
  PCA1weight[,j]=pca$rotation[,1];
  names(tmp)=1:length(tmp);
  #barplot(tmp,cex.names=1.25,xlab="PCA1 Loading",main=paste(gsub(".xlsx","",list_file[j]),"Original"))
  
  tmp=pca$rotation[,2];
  PCA2weight[,j]=pca$rotation[,2];
  names(tmp)=1:length(tmp); #第二主成分
  #barplot(tmp,cex.names=1.25,xlab="PCA2 Loading")
  #dev.off();

  autoplot(pca, data=A1, loadings=TRUE, loadings.colour='blue',
           loadings.label=TRUE, loadings.label.size=4,
           main=paste0("The biplot of station ",j," (",substr(list_file[j],start=4,stop=5),
                       ") in ",as.numeric(substr(list_file[j],start=6,stop=8))+1911))
  ggsave(paste0(substr(list_file[j],start=1,stop=9),"_Biplot.png"))
}

for(j in 1:length(list_file))
{
  j=9
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #去掉日期與溫度
  pca <- prcomp( A1, scale = FALSE) #建立pca
  cat(substr(list_file[j],start=1,stop=6),"\n",
      "PCA1累計解釋變異比例:",summary(pca)$importance[3,1],"\n",
      "PCA2累計解釋變異比例:",summary(pca)$importance[3,2],"\n")
}

###################################
PCACSV=function(Station,Order)
{
  Y=array(NA,nrow(list_A[[Station]]))
  if(Order==1)     {PCAW=PCA1weight; PCAWN="PC1.csv"}
  else if(Order==2){PCAW=PCA2weight; PCAWN="PC2.csv"}
  else             {PCAW=matrix(0,54,8);  PCAWN="error.csv"}
  for(Time in c(1:nrow(Y)))
  {
    temp1=as.numeric(list_A[[Station]][Time,-c(1,56)])
    temp2=as.numeric(PCAW[,Station])
    
    Y[Time]=temp1%*%temp2
  }
  Day=list_A[[Station]][,1]
  write.csv(cbind(Day,Y),paste0(substr(list_file,start=1,stop=2)[Station],"_",
                                substr(list_file,start=4,stop=14)[Station],PCAWN),
            row.names=F, quote=F)
}

for(station in 1:8){
  for(order in 1:2){PCACSV(Station=station,Order=order)}
}

A=NULL
for(j in 1:9)
{
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]
  pca <- prcomp( A1, scale = FALSE)
  A=rbind(A,
          rev(tail(order(abs(pca$rotation[,1])),5)),
          rev(tail(order(abs(pca$rotation[,2])),5)))
}
View(A)
write.csv(A,"PC12top5.csv",row.names=F, quote=F)
######

j=1
Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]) #去掉日期與溫度

r=3
res <- nmf(Data, rank=r)
par(mfrow=c(1,1))
basismap(res)
coefmap(res)
temp2=res@fit@H[1,]
names(temp2)=1:54
par(mfrow=c(r,1))
barplot(temp2,cex.names=.8,xlab="NMF Loading")


View(res@fit@W%*%res@fit@H)

res=nmf(matrix(data=c(1:6),nrow = 2,ncol = 3),1)
res@fit@W
res@fit@H
###
j=4
Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
r=4
res <- nmf(Data, rank=r)
par(mfrow=c(r,1))
for (base in 1:r)
{
  temp=res@fit@W[,base]
  names(temp)=1:54
  barplot(temp,cex.names=.8,xlab="NMF Loading")
}
name=list_file[j]
##################NMF#############
#######比較各rank下NMF的解釋變異
nmfev=nmfev1=matrix(NA,6,9)
for(j in 1:9)
{
  Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
  print(paste("Station ",j))
  for(r in 1:6)
  {
    res <- nmf(Data, rank=r)
    nmfev1[r,j]=1-sum((Data-res@fit@W%*%res@fit@H)^2)/sum((Data-mean(Data))^2)
    nmfev[r,j]=1-sum((Data-res@fit@W%*%res@fit@H)^2)/sum(Data^2)
    print(paste("Rank:",r," Explained Variance:",nmfev1[r,j]))
  }
}
###比較PCA的累計主成分解釋變異
pcaev=matrix(NA,6,9)
for(j in 1:9)
{
  Data=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]
  pca=prcomp(Data, scale = FALSE)
  pcaev[,j]=summary(pca)$importance[3,1:6]
}

##PCA NMF比較繪圖
Color=brewer.pal(9,"Set1")
par(mfrow=c(1,3))
plot(x=c(1:nrow(nmfev1)),y=nmfev1[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Rank",ylab="Explained Variance",main='(demean)The Explained Variance for Target V')
for(st in 2:9)
{lines(x=c(1:nrow(nmfev1)),y=nmfev1[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("萬華","土城","忠明","臺西","朴子","臺南","橋頭","小港","潮州"))

plot(x=c(1:nrow(nmfev)),y=nmfev[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Rank",ylab="Explained Variance",main='The Explained Variance for Target V')
for(st in 2:9)
{lines(x=c(1:nrow(nmfev)),y=nmfev[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("萬華","土城","忠明","臺西","朴子","臺南","橋頭","小港","潮州"))

plot(x=c(1:nrow(pcaev)),y=pcaev[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Principal Components",ylab="Cumulative Explained Variance",main='The Cumulative Explained Variance for Principal Components')
for(st in 2:9)
{lines(x=c(1:nrow(pcaev)),y=pcaev[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("萬華","土城","忠明","臺西","朴子","臺南","橋頭","小港","潮州"))
##########
set.seed(1)
for(j in 1:9)
{
  Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
  r=4
  res <- nmf(Data, rank=r)
  par(mfrow=c(r,1))
  for (base in 1:r)
  {
    temp=res@fit@W[,base]
    names(temp)=1:54
    barplot(temp,cex.names=.8,xlab="NMF Loading",main=paste("St",j))
  }

  ImportantPLOT=NULL
  for(i in 1:4)
  {
    basis=res@fit@W[,i]
    topbasis=basis[rev(tail(order(basis),10))]; names(topbasis)=rev(tail(order(basis),10))
    ImportantPLOT=c(ImportantPLOT,
                    as.numeric(names(topbasis[topbasis>((max(basis)-min(basis))/2)])))
  }
  ImportantPLOT=unique(sort(ImportantPLOT))
  print(paste("St",j))
  print(ImportantPLOT)
}

