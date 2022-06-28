rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr")

df <-readxl::read_xlsx("C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Clínica Quartier 404 - Thiago Ramalho/Clientes Dr.Tiago Ramalho.xlsx")
dfFinal <- data.frame()%>%
  mutate("Nome", "Nascimento", "Natural", "Profissao", "CPF",
         "RG","Email", "indicacao", "Telefone", "Logradouro", 
         "Bairro", "Cidade", "CEP")

linhas <- nrow(df)
colunas <- ncol(df)
l = 1
x = 0

while(l <= linhas){
  c = 1
  if(is.na(df[l,1]) == TRUE ){
    df[l,1] = "celula vazia"
  }
  
  if(is.na(df[l,2]) == TRUE ){
    df[l,2] = "celula vazia"
  }
  
  if(is.na(df[l,3]) == TRUE ){
    df[l,3] = "celula vazia"
  }
  
  if(str_detect(df[l,1],"celula vazia") == TRUE){
    x = x + 1
  }
  if(l > 1 && df[l-1, 1] == "celula vazia"){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,1] <- df[l,c]
    }else {
      dfFinal[x,1] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "Data de Nascimento:") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,2] <- df[l,c]
    }else {
      dfFinal[x,2] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "Natural:") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,3] <- df[l,c]
    }else {
      dfFinal[x,3] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "Profissão:") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,4] <- df[l,c]
    }else {
      dfFinal[x,4] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "CPF") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,5] <- df[l,c]
    }else {
      dfFinal[x,5] <- str_c(df[l,c], df[l,c + 1]) 
    }
  } 
  
  if(str_detect(df[l,c], "RG ") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,6] <- df[l,c]
    }else {
      dfFinal[x,6] <- str_c(df[l,c], df[l,c + 1]) 
    }
  } 
  
  if(str_detect(df[l,c], "E-mail:") == TRUE ||
     str_detect(df[l,c], "E-Mail") == TRUE ||
     str_detect(df[l,c], "email") == TRUE){
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,7] <- df[l,c]
    }else {
      dfFinal[x,7] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "Indicação") == TRUE ||
     str_detect(df[l,c], "Indicado") == TRUE ||
     str_detect(df[l,c], "Indicada") == TRUE ){
    if(is.na(df[l,c + 1])|| df[l,c + 1] == "celula vazia"){
      dfFinal[x,8] <- df[l,c]
    }else {
      dfFinal[x,8] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c + 1], "Indicação") == TRUE ||
     str_detect(df[l,c + 1], "Indicado") == TRUE ||
     str_detect(df[l,c + 1], "Indicada") == TRUE ){
    if(is.na(df[l,c + 2]) || df[l,c + 2] == "celula vazia"){
      dfFinal[x,8] <- df[l,c + 1]
    }else {
      dfFinal[x,8] <- str_c(df[l,c], df[l,c + 2]) 
    }
  }
  
  if(str_detect(df[l,c], "Telefone ") == TRUE ||
     str_detect(df[l,c], "Tel ") == TRUE ||
     str_detect(df[l,c], "TELEFONE ") == TRUE ||
     str_detect(df[l,c], "TEL ") == TRUE ||
     str_detect(df[l,c], "cel.") == TRUE ||
     str_detect(df[l,c], "cel ") == TRUE ||
     str_detect(df[l,c], "cel.:") == TRUE ||
     str_detect(df[l,c], "celula vazia") == FALSE &&
     str_detect(df[l,c], "Indica") == FALSE &&
     str_detect(df[l,c], "Estrada dos Coqueiros") == FALSE &&
     str_detect(df[l,c], "E-mail") == FALSE && 
     str_detect(df[l,c], "email") == FALSE && 
     str_detect(df[l,c], "Filha:") == FALSE &&
     str_detect(df[l,c], "CEP") == FALSE &&
     str_detect(df[l,c], "CPF") == FALSE &&
     str_detect(df[l,c], "Rua") == FALSE &&
     str_detect(df[l,c], "Convênio") == FALSE &&
     str_detect(df[l,c], "Logradouro") == FALSE &&
     str_detect(df[l,c], "Endereço") == FALSE &&
     str_detect(df[l,c], "Mãe") == FALSE &&
     str_detect(df[l,c], "Nascimento") == FALSE &&
     str_detect(df[l,c], "Profissão") == FALSE &&
     str_detect(df[l,c], "Natural") == FALSE &&
     str_detect(df[l,c], "Nome") == FALSE &&
     str_detect(df[l,c], "Procedimento") == FALSE &&
     str_detect(df[l,c], "procedimento") == FALSE ){	
    if(is.na(df[l,c + 1]) || df[l,c + 1] == "celula vazia"){
      dfFinal[x,9] <- df[l,c]
    }else {
      dfFinal[x,9] <- str_c(df[l,c], df[l,c + 1]) 
    }
  }
  
  if(str_detect(df[l,c], "Endereço") == TRUE &&
     str_detect(df[l + 1,c], "Logradouro/") == TRUE){
      dfFinal[x,10] <- df[l + 2,c]
      dfFinal[x,11] <- df[l + 2,c + 1]
      dfFinal[x,12] <- df[l + 2,c + 2]
      dfFinal[x,13] <- df[l + 2,c + 3]
  }
  
  if(str_detect(df[l,c], "Endereço") == TRUE &&
     nchar(df[l,c]) > 13){
    dfFinal[x,10] <- df[l,c]
  }
  
  if(str_detect(df[l,c], "Endereço") == TRUE &&
     str_detect(df[l + 1,c], "Logradouro/") == FALSE &&
     str_detect(df[l + 1,c], "CEP") == FALSE &&
     nchar(df[l,c]) < 10){
    dfFinal[x,10] <- df[l + 1,c]
    dfFinal[x,11] <- df[l + 1,c + 1]
    dfFinal[x,12] <- df[l + 1,c + 2]
    dfFinal[x,13] <- df[l + 1,c + 3]
  }
  
  l = l +1
}

writexl::write_xlsx(dfFinal, "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Clínica Quartier 404 - Thiago Ramalho/Dr.Tiago Ramalho - final.xlsx")


