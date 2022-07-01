rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr", "readxl")

df <- read_xlsx("C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/MDMED/Carla Mazoler - P.xlsx")
df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)
dfFinal <- data.frame()%>%
  mutate("Nome", "Cadastro", "Código", "Endereço", "Bairro", "Cidade", "Estado", "Cep", "Fone",
         "Data de nascimento", "Sexo", "Estado civil",
         "Convenio", "Profissão", "ultima consulta")

colunas <- ncol(df)
linhas <- nrow(df)


l = 1
c = 1
x = 0

while(l <= linhas){
  
  if(str_detect(df[l,c], "Nome:") == TRUE &&
     str_detect(df[l,c], "Cadastro") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Cadastro")
    x = x + 1
    dfFinal[x,1] <- ateColuna3[[1]][1]
    dfFinal[x,2] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Código:")
    dfFinal[x,3] <- ateColuna3[[1]][2]
  }
  
  if(str_detect(df[l,c], "Endereço") == TRUE &&
     str_detect(df[l,c], "Bairro") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Bairro:")
    dfFinal[x,4] <- ateColuna3[[1]][1]
    dfFinal[x,5] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Cidade:")
    dfFinal[x,6] <- ateColuna3[[1]][2]
  }
  if(str_detect(df[l,c], "Estado") == TRUE &&
     str_detect(df[l,c], "Cep") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Cep:")
    dfFinal[x,7] <- ateColuna3[[1]][1]
    dfFinal[x,8] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Fone:")
    dfFinal[x,9] <- ateColuna3[[1]][2]
  }
  if(str_detect(df[l,c], "Data Nasc:") == TRUE &&
     str_detect(df[l,c], "Sexo:") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Sexo:")
    dfFinal[x,10] <- ateColuna3[[1]][1]
    dfFinal[x,11] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Est. Civi")
    dfFinal[x,12] <- ateColuna3[[1]][2]
  }
  if(str_detect(df[l,c], "Data Nasc:") == TRUE &&
     str_detect(df[l,c], "Sexo:") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Sexo:")
    dfFinal[x,10] <- ateColuna3[[1]][1]
    dfFinal[x,11] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Est. Civi")
    dfFinal[x,12] <- ateColuna3[[1]][2]
  }
  if(str_detect(df[l,c], "Convênio") == TRUE &&
     str_detect(df[l,c], "Profissão:") == TRUE){
    ateColuna3 <- strsplit(df[l,c], split = "Profissão:")
    dfFinal[x,13] <- ateColuna3[[1]][1]
    dfFinal[x,14] <- ateColuna3[[1]][2]
    ateColuna3 <- strsplit(df[l,c], split = "Ult. Cons:")
    dfFinal[x,15] <- ateColuna3[[1]][2]
  }

  l = l + 1
}


writexl::write_xlsx(dfFinal,"C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/MDMED/Carla Mozaler - P - Final.xlsx")
