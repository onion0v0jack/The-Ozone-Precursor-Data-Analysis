rm(list=ls())  #�M���Ҧ��Ȧs
graphics.off() #�M���Ϥ������
getwd()        #��ܳo��r�Ҧb�����|

Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_144') 
require(ggplot2)
library(ggfortify)
library(readxl)
require(RColorBrewer)
require(NMF)
setwd("C:/Users/user/�U��/���s����/Meeting/�פ�/����/106�~�e�X�����/�饭�����")
list_file=list.files(pattern ="*.xlsx")#[-9]; #�Ǧ^���w���|�����ɮצW�٦r���V�q�Τl�ؿ��C

list_A=list(); #���ɦW�blist_file�����ɮ׻`�����@��list
for(i in 1:length(list_file))
{
  list_A[[i]]=readxl::read_excel(list_file[[i]],col_names = TRUE)
}

which(is.na(list_A[[9]]),arr.ind =T) #�ݬO�_��NA

###########�H�U���ζ]#################
mod_A=list()
for(j in 1:length(list_file))
{
  j=1
  A=list_A[[j]]
  mod_A[[j]]=A;
  for(i in 2:ncol(A))#�]���C�@�Ӹ�ƪ��Ĥ@��col���N�����
  {
    marfit=arima(A[,i],order = c(1,0,0) ); #A[,i]���ܨC�@�Ӹ�ƪ����e�X��(�̫�@�Ӭ��ū�)
    mod_A[[j]][,i]=marfit$residuals; #�Q��AR(1)�t�A�C�@�ӫe�X�����饭���ƭȡA�åB�N��ݮt�s�bmod_A��
  }
}#mod_A���Ҧ���ƪ�size���ܡA����������]�S���ܡA�u����L�ƭȴݮt�Ӥw

#Loading(�ϥΨ�ݮt�A�]�N�Omod_A�����)
for(j in 1:length(list_file))
{
  A1=mod_A[[j]][,-1*c(1,ncol(mod_A[[j]]))];#������Ƥ��������ūפ��p
  pca <- prcomp( A1, scale = FALSE) 
  
  png(filename = paste(list_file[j],"loading",".png",sep=""),
      width = 2048, height = 1024) #�إߤ@��png�ó]�w�ɦW�A���Ϫ����e�ٻݭn�U���M�w
  par(mfrow=c(2,1))
  tmp=pca$rotation[,1]; #�Ĥ@�D����(���V�q���U�e�X�����v��)
  names(tmp)=1:length(tmp);
  barplot(tmp,cex.names=1.25,xlab="PCA1 Loading",main=paste(gsub(".xlsx","",list_file[j]),"De-Time Series"))
  
  tmp=pca$rotation[,2]; #�ĤG�D����
  names(tmp)=1:length(tmp);
  barplot(tmp,cex.names=1.25,xlab="PCA2 Loading")
  dev.off();
  
  #��Ĥ@�D������ĤG�D�����إߤ@�ӥ����A�ç�U�e�X��������b��D���������q�A��ܥX��
  autoplot(pca, data = A1,
           loadings = TRUE, loadings.colour = 'blue',
           loadings.label = TRUE, loadings.label.size = 4,main=paste(gsub(".xlsx","",list_file[j]),"De-Time Series"))
  ggsave( paste(list_file[j],"loading_scores",".png",sep="")) #�s��
}
###########�H�W���ζ]#################

#original(�ϥέ쥻��� �]�N�Olist_A)
PCA1weight=PCA2weight=as.data.frame(matrix(NA,54,9))
colnames(PCA1weight)=colnames(PCA2weight)=c("S1","S2","S3","S4","S5","S6","S7","S8","S9")
for(j in 1:length(list_file))
{
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #�h������P�ū�
  pca <- prcomp( A1, scale = FALSE) #�إ�pca
  
  #png(filename = paste(list_file[j],"loading_original",".png",sep=""),
  #    width = 2048, height = 1024)
  #par(mfrow=c(2,1))
  tmp=pca$rotation[,1]; #�Ĥ@�D����
  PCA1weight[,j]=pca$rotation[,1];
  names(tmp)=1:length(tmp);
  #barplot(tmp,cex.names=1.25,xlab="PCA1 Loading",main=paste(gsub(".xlsx","",list_file[j]),"Original"))
  
  tmp=pca$rotation[,2];
  PCA2weight[,j]=pca$rotation[,2];
  names(tmp)=1:length(tmp); #�ĤG�D����
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
  A1=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]; #�h������P�ū�
  pca <- prcomp( A1, scale = FALSE) #�إ�pca
  cat(substr(list_file[j],start=1,stop=6),"\n",
      "PCA1�֭p�����ܲ����:",summary(pca)$importance[3,1],"\n",
      "PCA2�֭p�����ܲ����:",summary(pca)$importance[3,2],"\n")
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
Data=t(list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]) #�h������P�ū�

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
#######����Urank�UNMF�������ܲ�
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
###���PCA���֭p�D���������ܲ�
pcaev=matrix(NA,6,9)
for(j in 1:9)
{
  Data=list_A[[j]][,-1*c(1,ncol(list_A[[j]]))]
  pca=prcomp(Data, scale = FALSE)
  pcaev[,j]=summary(pca)$importance[3,1:6]
}

##PCA NMF���ø��
Color=brewer.pal(9,"Set1")
par(mfrow=c(1,3))
plot(x=c(1:nrow(nmfev1)),y=nmfev1[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Rank",ylab="Explained Variance",main='(demean)The Explained Variance for Target V')
for(st in 2:9)
{lines(x=c(1:nrow(nmfev1)),y=nmfev1[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("�U��","�g��","����","�O��","���l","�O�n","���Y","�p��","��{"))

plot(x=c(1:nrow(nmfev)),y=nmfev[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Rank",ylab="Explained Variance",main='The Explained Variance for Target V')
for(st in 2:9)
{lines(x=c(1:nrow(nmfev)),y=nmfev[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("�U��","�g��","����","�O��","���l","�O�n","���Y","�p��","��{"))

plot(x=c(1:nrow(pcaev)),y=pcaev[,1],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[1],
     cex.lab=1.5,xlab="Principal Components",ylab="Cumulative Explained Variance",main='The Cumulative Explained Variance for Principal Components')
for(st in 2:9)
{lines(x=c(1:nrow(pcaev)),y=pcaev[,st],type="o",pch=19,lwd=2,ylim=c(0.4,1),col=Color[st])}
abline(0.8,0,lty=2,lwd=2)
legend("bottomright",col=Color,lwd=5,cex=1.5,
       legend=c("�U��","�g��","����","�O��","���l","�O�n","���Y","�p��","��{"))
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
