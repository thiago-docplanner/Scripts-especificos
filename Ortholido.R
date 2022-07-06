rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr", "readxl")

df <- read.csv("C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/PDF ORTHOLIDO/tabula-CADASTRO PACIENTES ORTHOLIDO PDF.csv")

df <- data.frame(lapply(df, as.character), stringsAsFactors=FALSE)
dfFinal <- data.frame()%>%
  mutate("Nome", "Data de nascimento", "Endereço", "Telefone", 
         "Email", "data da consulta")

colunas <- ncol(df)
linhas <- nrow(df)


l = 1
c = 1
x = 0
x = 0
while(l <= linhas){
  
  if(str_detect(df[l,c], "1") == TRUE ||
     str_detect(df[l,c], "2") == TRUE ||
     str_detect(df[l,c], "3") == TRUE ||
     str_detect(df[l,c], "4") == TRUE ||
     str_detect(df[l,c], "5") == TRUE ||
     str_detect(df[l,c], "6") == TRUE ||
     str_detect(df[l,c], "7") == TRUE ||
     str_detect(df[l,c], "8") == TRUE ||
     str_detect(df[l,c], "9") == TRUE){
    if(str_detect(df[l,c], '\\)') == TRUE){
      x = x + 1
      dfFinal[x,1] <- df[l,c]
      if(str_detect(df[l + 1 ,c], 'ascimento') == TRUE){
        dfFinal[x, 2] <- df[l + 1 ,c]
      }
      else if(str_detect(df[l + 2 ,c], 'ascimento') == TRUE){
        dfFinal[x, 2] <- df[l + 2 ,c]
      }
      else if(str_detect(df[l + 3 ,c], 'ascimento') == TRUE){
        dfFinal[x, 2] <- df[l + 3 ,c]
      }
      if(str_detect(df[l + 1 ,c], 'ndereÃ§') == TRUE ||
         str_detect(df[l + 1 ,c], 'End: ') == TRUE){
        dfFinal[x, 3] <- df[l + 1 ,c]
        if(str_detect(df[l + 2 ,c], 'ascimento') == FALSE &&
           str_detect(df[l + 2 ,c], 'Tel') == FALSE){
          dfFinal[x, 3] <- str_c(dfFinal[x, 3], " - ", df[l + 2 ,c])
        }
      }
      else if(str_detect(df[l + 2 ,c], 'ndereÃ§') == TRUE ||
              str_detect(df[l + 2 ,c], 'End: ') == TRUE){
        dfFinal[x, 3] <- df[l + 2 ,c]
        
        if(str_detect(df[l + 3 ,c], 'ascimento') == FALSE &&
           str_detect(df[l + 3 ,c], 'Tel') == FALSE){
          dfFinal[x, 3] <- str_c(dfFinal[x, 3], " - ", df[l + 2 ,c])
        }
      }
    }
  }
  if(str_detect(df[l,c], "Tel:") == TRUE ||
     str_detect(df[l,c], "Telefone") == TRUE){
    dfFinal[x, 4] <- df[l,c]
  }
  if(str_detect(df[l,c], "Email:") == TRUE ||
     str_detect(df[l,c], "email:") == TRUE){
    dfFinal[x, 5] <- df[l,c]
  }
  if(str_detect(df[l,c], "Consulta") == TRUE ||
     str_detect(df[l,c], "consulta") == TRUE){
    dfFinal[x, 6] <- df[l,c]
  }

  l = l + 1
}


writexl::write_xlsx(dfFinal,"C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/PDF ORTHOLIDO/tabula-CADASTRO PACIENTES ORTHOLIDO PDF-Final.xlsx")
