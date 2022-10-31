library(readxl)
library(dplyr)
library(openxlsx)


setwd("C:/Users/46392/Desktop/Issues/ISSUE 290655")


X= "META1" #Troca o nome conforme a meta

CD_Geral= read_excel("META 1 CD.xlsx") %>% 
  arrange(`Id Pergunta`)  #Troca o nome conforme a meta

DEV_Geral = read_excel("META 1 DEV.xlsx") %>% 
  arrange(`Id Pergunta`)  #Troca o nome conforme a meta

CD_Geral$`Sum([Quantidade Processos])` = NULL
DEV_Geral$`Sum([Quantidade Processos])` = NULL

CD_Geral$`#` = NULL
DEV_Geral$`#` = NULL



CD = list()

for (i in 1:length(unique(CD_Geral$`Id Pergunta`))){
  
  CD[[i]] = CD_Geral %>% 
    filter(`Id Pergunta` == unique(`Id Pergunta`)[[i]])
  
  i = i+1
  
}



DEV = list()

for (i in 1:length(unique(DEV_Geral$`Id Pergunta`))){
  
  DEV[[i]] = DEV_Geral %>% 
    filter(`Id Pergunta` == unique(`Id Pergunta`)[[i]])
  
  i = i+1
  
}

CD_Geral_DEV = list() #META1_Tem_Dados_Nao_DEV 



for (i in 1:length(unique(CD))){
  
  CD_Geral_DEV[[i]] = anti_join((as.data.frame(CD[[i]])), as.data.frame(DEV[[i]]) , "NumeroUnico")
  
  i = i+1
  
}

Nomes_CD = NULL

for  (i in 1:length(CD_Geral_DEV )){
  
  assign(paste0(X,"_CD_DEV_",unique(DEV_Geral$`Id Pergunta`)[i]), as.data.frame(CD_Geral_DEV[[i]]))
  
  Nomes_CD[i]= paste0(X,"_CD_DEV_",unique(DEV_Geral$`Id Pergunta`)[i])
  
}


DEV_Geral_CD= list() #META1_Tem_Dev_Nao_Dados 


for (i in 1:length(unique(CD))){
  
  DEV_Geral_CD[[i]] = anti_join(as.data.frame(DEV[[i]]),(as.data.frame(CD[[i]])) , "NumeroUnico")
  
  i = i+1
  
}

Nomes_DEV = NULL

for  (i in 1:length(DEV_Geral_CD)){
  
  assign(paste0(X,"_DEV_CD_",unique(DEV_Geral$`Id Pergunta`)[i]), as.data.frame(DEV_Geral_CD[[i]]))
  
  Nomes_DEV[i] = paste0(X,"_DEV_CD_",unique(DEV_Geral$`Id Pergunta`)[i])
  
  
}


#CD_DEV Tem no ciência de dados e não no dev
#DEV_CD Tem no dev e não no ciência de dados



setwd("C:/Users/46392/Desktop/Issues/ISSUE 290655/Comparação")


WB_CD <- createWorkbook()
for (i in 1: length(Nomes_CD)){
  addWorksheet(WB_CD,Nomes_CD[i])
  writeData(WB_CD, i, get(Nomes_CD[i]))
  
}

WB_DEV <- createWorkbook()
for (i in 1: length(Nomes_DEV)){
  addWorksheet(WB_DEV,Nomes_DEV[i])
  writeData(WB_DEV, i, get(Nomes_DEV[i]))
  
}

saveWorkbook(WB_CD, paste0(X,"_CD_DEV",".xlsx"))
saveWorkbook(WB_DEV, paste0(X,"_DEV_CD",".xlsx"))

remove(CD,DEV,CD_Geral,CD_Geral_DEV, DEV_Geral,DEV_Geral_CD)



#for (i in 1: length(Nomes_DEV)){

#  write.xlsx(get(Nomes_DEV[i]),paste0(Nomes_DEV[i],".xlsx"),sheetName= Nomes_DEV[i])

#}


#for (i in 1: length(Nomes_CD)){

#  write.xlsx(get(Nomes_CD[i]),paste0(Nomes_CD[i],".xlsx"),sheetName= Nomes_CD[i])

#}



