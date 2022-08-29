rm(list = ls())
pacman::p_load('readxl','stringr', 'dplyr')

dir <- "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/em andamento/ExpDadosACOMP_CLINICO/"
diretorio <- str_c(dir,"input/")

arquivos <- list.files(diretorio, pattern = ".*.txt")

dfFinal <- data.frame() %>%
  mutate("arquivo", "texto")
l = 1

for (arquivo in arquivos){
  dfFinal[l,1] <- arquivo
  txt <- read.table(str_c(diretorio,arquivo),
                    header = FALSE, skipNul = TRUE, fill= TRUE, sep="\t")
  y = 1
  linhas <- nrow(txt)
  colunas <- ncol(txt)
  while(y <= linhas){
    if(y == 1){
      dfFinal[l,2] <-  txt[y,1]
      if(is.na(txt[y,2]) == FALSE && colunas > 1){
        dfFinal[l,2] <- str_c(dfFinal[l,2], " ", txt[y,1])
      }
    }else if(txt[y,1] != ""){
      dfFinal[l,2] <- str_c(dfFinal[l,2], "\n", txt[y,1])
    }else {
      dfFinal[l,2] <- str_c(dfFinal[l,2], "\n", txt[y,2])
    }
    y = y + 1
  }
  l = l + 1
  print(arquivo)
}
writexl::write_xlsx(dfFinal,str_c(dir, "final.xlsx"))