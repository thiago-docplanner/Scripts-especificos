pacman::p_load('readxl','stringr', 'dplyr')

diretorio <- "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/IMedicina/input/"

diretorioFinal <- "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/IMedicina/output/"

arquivo <- "bkp_imed_16-03-2022 (1).xlsx"


arquivoInicial <- str_c(diretorio, arquivo)
  
df <- readxl::read_xlsx(arquivoInicial)
  
dfFinal <- data.frame()%>%
  mutate("ID", "Nome","evento", "Inicio", "Anamnese")
  
linhas <- nrow(df)
colunas <- ncol(df)

l = 0
y = 0
  
while(y<=linhas){
  y = y+1
  x = 3
    
  while(x<=colunas){
    c = 1
    if(is.na(df[y,x]) == FALSE){
      l = l + 1
      dfFinal[l,c] <- df[y,1]
      c = c + 1
      dfFinal[l,c] <- df[y,2]
      # Fazer split do atendimento
      splitInicio <-c(str_split(df[y,x], "inicio:"))
      c = c + 1
      dfFinal[l,c] <- splitInicio[[1]][1]
        
        
      c = c + 1
      splitFim <-c(str_split(splitInicio[[1]][2], "fim:"))
      dfFinal[l,c] <- splitFim[[1]][1]
        
        
      c = c + 1
      splitAnamnese <-c(str_split(splitInicio[[1]][2], "anamnese:"))
      dfFinal[l,c] <- splitAnamnese[[1]][2]
    }
    x = x +1
  }
}
  
  
arquivo_final <- str_c(diretorioFinal, arquivo)
  
writexl::write_xlsx(dfFinal,arquivo_final)
