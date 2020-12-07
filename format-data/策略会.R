setwd("C:\\Users\\GL\\Desktop\\2019年中期策略会\\代码")
library(readxl)
data=read_excel("test.xlsx",sheet = 1)
company=data[,c(3,5)]
for(i in 6:ncol(data)){
  aa=data[,c(3,i)]
  colnames(aa)=colnames(company)
  company=rbind(company,aa)
  print(i)
}  
#以上为改为：人，公司这样的形式

company=na.omit(data.frame("company"=unique(company[,2]),stringsAsFactors = F))
result=list()
num=c()
for(i in 1:nrow(company)){
  aa=data[which(data[,5]==company[i,1]),]
  for(j in 6:ncol(data)){
    aa=rbind(aa,data[which(data[,j]==company[i,1]),])
  }
  sub=data.frame()
  for(j in 1:nrow(aa)){
    sub[1,1]=NA
    sub[1,2]=NA
    sub[j,3]=aa[j,1]
    sub[j,4]=paste(aa[j,2],"-",aa[j,3])
    sub[j,5]=aa[j,4]
  }
  sub=sub[order(sub[,5]),]
  sub[1,1]=company[i,1]
  sub[1,2]=nrow(aa)
  num=c(num,nrow(aa))
  result[[i]]=sub
  print(i)
}
conclusion=result[[order(num,decreasing = T)[1]]]
for(i in order(num,decreasing = T)[2:length(num)]){
  conclusion=rbind(conclusion,result[[i]])
  print(i)
}
colnames(conclusion)=c("公司","人数","销售","机构-姓名","分类")
write.csv(conclusion,"上市公司交流.csv")
