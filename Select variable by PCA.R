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
{list_A[[i]]=readxl::read_excel(list_file[[i]],col_names = TRUE)}

which(is.na(list_A[[9]]),arr.ind =T) #看是否有NA
  
#original(使用原本資料 也就是list_A)
PCA1weight=PCA2weight=as.data.frame(matrix(NA,54,9))
colnames(PCA1weight)=colnames(PCA2weight)=c("S1","S2","S3","S4","S5","S6","S7","S8","S9")
for(j in 1:length(list_file))
{
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #去掉日期與溫度
  pca <- prcomp( A1, scale = FALSE) #建立pca
  
  png(filename=paste(substr(list_file[j],start=1,stop=5),"PC12.png",sep=""),
      width=2048, height=1024)
  par(mfrow=c(2,1))
  par(mar = c(5.1,5,4.1,0)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)par(mfrow=c(2,1))
  tmp=pca$rotation[,1]; #第一主成分
  PCA1weight[,j]=pca$rotation[,1];
  names(tmp)=1:length(tmp);
  barplot(tmp,cex.names=1.4,cex.lab=1.8,cex.main=3,
          ylab="PC1 Loading", xlab="Number of precursor",
          main=paste0("Principal component of Station ",j," (",paste(substr(list_file[j],start=4,stop=5)),
                      ") in ",as.numeric(paste(substr(list_file[j],start=6,stop=8)))+1911))
  tmp=pca$rotation[,2];
  PCA2weight[,j]=pca$rotation[,2];
  names(tmp)=1:length(tmp); #第二主成分
  barplot(tmp,cex.names=1.4,cex.lab=1.8,cex.main=3,
          ylab="PC2 Loading", xlab="Number of precursor")
  dev.off();
}

#前二主成分累積解釋變異
for(j in 1:length(list_file))
{
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #去掉日期與溫度
  pca <- prcomp( A1, scale = FALSE) #建立pca
  cat(substr(list_file[j],start=1,stop=6),"\n",
      "PC1解釋變異比例:",summary(pca)$importance[2,1],"\n",
      "PC2解釋變異比例:",summary(pca)$importance[2,2],"\n",
      "PC3解釋變異比例:",summary(pca)$importance[2,3],"\n")
}

#看各站各主成分權重排序與選取結果(絕對值排序取最大差分之前)
{
  j=3
  PC=2
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]
  pca=prcomp(A1, scale=FALSE)
  tmp=pca$rotation[,PC]
  names(tmp)=1:length(tmp)
  barplot(tmp)
  
  OrderPLOT=rev(order(abs(tmp)))#排名由大到小
  OrderAbs=abs(tmp[OrderPLOT])
  barplot(OrderAbs)
  
  OrderPLOT[1:order(diff(array(OrderAbs)))[1]]
}

#將前三主成分取絕對值，以解釋變異作為權重加權後排序取前驅物(最大差分之前)
{
  for(j in 1:9){
    A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]
    pca=prcomp(A1, scale=FALSE)
    PC1v=summary(pca)$importance[2,1]; PC1=abs(pca$rotation[,1]); names(PC1)=1:length(PC1)
    PC2v=summary(pca)$importance[2,2]; PC2=abs(pca$rotation[,2]); names(PC2)=1:length(PC2)
    PC3v=summary(pca)$importance[2,3]; PC3=abs(pca$rotation[,3]); names(PC3)=1:length(PC3)
    PCweight=PC1v*PC1+PC2v*PC2+PC3v*PC3
    barplot(PCweight)
    
    OrderPLOT=rev(order(abs(PCweight)))#排名由大到小
    Order=PCweight[OrderPLOT]
    barplot(Order)
    
    cat(substr(list_file[j],start=1,stop=5),": ",
        OrderPLOT[1:order(diff(array(Order)))[1]],"\n")
    #OrderPLOT[1:order(diff(array(Order)))[1]]
  }
}


###################################
#將資料丟到前二主成分，每站分別產出兩個時間序列
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
                                substr(list_file,start=4,stop=9)[Station],PCAWN),
            row.names=F, quote=F)
}

for(station in 1:9){
  for(order in 1:2){PCACSV(Station=station,Order=order)}
}

########
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

#######比較各rank下NMF的解釋變異
nmfev=nmfev1=matrix(NA,6,9)
for(j in 1:9)
{
  Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))])
  print(paste("Station ",j))
  for(r in 1:6)
  {
    res <- nmf(Data, rank=r)
    PCS=NULL; for(pcs in 1:54){temp=Data[pcs,]; PCS[pcs]=sum((temp-mean(temp))^2)}
    nmfev1[r,j]=1-sum((Data-res@fit@W%*%res@fit@H)^2)/sum((Data-mean(Data))^2)
    nmfev[r,j]=1-sum((Data-res@fit@W%*%res@fit@H)^2)/sum(PCS)
    print(paste("Rank:",r," Explained Variance:",nmfev1[r,j]))
    print(paste("Rank:",r," Explained Variance:",nmfev[r,j]))
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
#三個圖一起
{
  Color=brewer.pal(9,"Set1")
  png(filename="A.png", width=2048, height=1024)
  par(mfrow=c(1,3))
  par(mar = c(5.1,6,4.1,3)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  plot(x=c(1:nrow(nmfev1)),y=nmfev1[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Rank",ylab="Explained variance",main='(demean)The Explained Variance NMF')
  for(st in 2:9)
  {lines(x=c(1:nrow(nmfev1)),y=nmfev1[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  
  plot(x=c(1:nrow(nmfev)),y=nmfev[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Rank",ylab="Explained variance",main='The Explained Variance for NMF')
  for(st in 2:9)
  {lines(x=c(1:nrow(nmfev)),y=nmfev[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  
  plot(x=c(1:nrow(pcaev)),y=pcaev[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Principal component",ylab="Cumulative explained variance",main='The Cumulative Explained Variance for PC')
  for(st in 2:9)
  {lines(x=c(1:nrow(pcaev)),y=pcaev[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  dev.off();
}
#兩個圖一起
{
  Color=brewer.pal(9,"Set1")
  png(filename="B.png", width=2048, height=1024)
  par(mfrow=c(1,2))
  par(mar = c(5.1,6,4.1,3)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  plot(x=c(1:nrow(nmfev1)),y=nmfev1[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Rank",ylab="Explained variance",main='The Explained Variance NMF')
  for(st in 2:9)
  {lines(x=c(1:nrow(nmfev1)),y=nmfev1[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  
  plot(x=c(1:nrow(pcaev)),y=pcaev[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Principal component",ylab="Cumulative explained variance",main='The Cumulative Explained Variance for PC')
  for(st in 2:9)
  {lines(x=c(1:nrow(pcaev)),y=pcaev[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  dev.off();
}

##########


{
  Color=brewer.pal(9,"Set1")
  png(filename="B.png", width=1600, height=1024)
  par(mfrow=c(1,1))
  par(mar = c(5.1,6,4.1,3)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  
  plot(x=c(1:nrow(nmfev)),y=nmfev[,1],type="o",pch=19,lwd=3,ylim=c(0.2,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Rank",ylab="Explained variance",main='The explained variance for NMF')
  for(st in 2:9)
  {lines(x=c(1:nrow(nmfev)),y=nmfev[,st],type="o",pch=19,lwd=3,ylim=c(0.3,1),col=Color[st])}
  abline(h=c(0.8,0.9),lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  dev.off();
}


{
  Color=brewer.pal(9,"Set1")
  png(filename="C.png", width=1600, height=1024)
  par(mfrow=c(1,1))
  par(mar = c(5.1,6,4.1,3)) #設定圖形空白邊界行數下左上右，Default=c(5.1, 4.1, 4.1, 2.1)
  
  plot(x=c(1:nrow(pcaev)),y=pcaev[,1],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[1],
       cex.axis=2.5,cex.lab=2.75,cex.main=3,
       xlab="Principal component",ylab="Cumulative explained variance",main='The cumulative explained variance for PC')
  for(st in 2:9)
  {lines(x=c(1:nrow(pcaev)),y=pcaev[,st],type="o",pch=19,lwd=3,ylim=c(0.4,1),col=Color[st])}
  abline(0.8,0,lty=2,lwd=2)
  legend("bottomright",col=Color,lwd=5,cex=3,
         legend=c("Station 1 (萬華)","Station 2 (土城)","Station 3 (忠明)",
                  "Station 4 (臺西)","Station 5 (朴子)","Station 6 (臺南)",
                  "Station 7 (橋頭)","Station 8 (小港)","Station 9 (潮州)"))
  dev.off();
}