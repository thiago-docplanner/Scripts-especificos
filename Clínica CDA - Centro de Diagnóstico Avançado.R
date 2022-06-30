rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr", "readxl")

df <- read_xlsx("C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Clínica CDA - Centro de Diagnóstico Avançado/arquivos/1.xlsx")
df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)
dfFInal <- data.frame()%>%
  mutate("Nome", "Sexo", "Data de Cadastro", "Prontuario", "Data de nascimento", "Cor", "Convênio", 
         "Endereço","Estado civil", "Plano/Empresa", "Bairro", "Profissão", "No MAT. / CARTEIRA:", 
         "Cidade", "Identidade","Data de validade", "CEP", "CPF", "Titular", "Telefone1", "Telefone2",
         "Telefone3","Telefone4", "Email", "N° cartão de saúde")

colunas <- ncol(df)
linhas <- nrow(df)
l = 1
x = 0

while (l<=linhas) {
  c = 1
  while (c <= colunas) {
    if(is.na(df[l,c])){
      df[l,c] = ""
    }
    c = c + 1
  }
  l = l + 1
}

l = 1
c = 1

while(l <= linhas){
  c = 1
  while(c <= colunas){
    if(str_detect(df[l,c], "NOME:") == TRUE){
      x = x + 1
      y = 1
      dfFInal[x,y] <- df[l,2]
    } 
    else if(str_detect(df[l,c], "SEXO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "DATA DE CADASTRO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "PRONTUÁRIO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "DT. NASCIMENTO:") == TRUE &&
            nchar(df[l,c]) > 21){
      y = y + 1
      dfFInal[x,y] <- df[l,c]
    }
    else if(str_detect(df[l,c], "DT. NASCIMENTO:") == TRUE &&
            nchar(df[l,c]) <= 21){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "COR:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "CONVÊNIO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "ENDEREÇO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "ESTADO CIVIL:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "PLANO / EMPRESA:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "BAIRRO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "PROFISSÃO:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "No MAT. / CARTEIRA:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "CIDADE:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "IDENTIDADE:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "DATA VALIDADE:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "CEP:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "CPF:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "TITULAR:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "Celular") == TRUE ||
            str_detect(df[l,c], "WhatsApp") == TRUE ||
            str_detect(df[l,c], "Residencial") == TRUE ||
            str_detect(df[l,c], "Comercial") == TRUE ||
            str_detect(df[l,c], "Telefone") == TRUE ||
            df[l,c] == ":"){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "E-MAIL:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    else if(str_detect(df[l,c], "No CARTÃO NAC. DE SAÚDE:") == TRUE){
      y = y + 1
      dfFInal[x,y] <- df[l,c + 1]
    }
    c = c + 1
  }
  l = l + 1
}


writexl::write_xlsx(dfFInal,"C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Clínica CDA - Centro de Diagnóstico Avançado/arquivos/1 - FInal.xlsx")