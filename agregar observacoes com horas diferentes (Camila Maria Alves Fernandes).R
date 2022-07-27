rm(list = ls())

pacman::p_load('dplyr','purrr','readxl','stringr')

df <-readxl::read_xlsx('C:/Users/thiago.oliveira/Downloads/AG 1.xlsx')
df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)
#### Tratamento ####


colunas <- ncol(df)
linhas <- nrow(df)
l = 1
x = 0

while (l<=linhas) {
  c = 1
  while (c <= colunas) {
    if(is.na(df[l,c])){
      df[l,c] = "celula vazia"
    }
    c = c + 1
  }
  l = l + 1
}

l = 1

while(l <= linhas){
  x = l + 1
  print(df[l,1])
  while(x <= linhas) {
    if(df[l,1] == df[x,1] ){
      if(df[l,3] == df[x,3] &&
         df[l,7] == df[x,7]){
        df <- slice(df, -x)
        linhas <- nrow(df)
      }
    } else if(df[l,1] != df[x,1]){
      x = linhas + 1
    }
    x = x + 1
  }
  l = l + 1
}

df2 <- df %>%
  aggregate(. ~  ID.Paciente+ Paciente+  Data  +
              Service.ID + Scheduled.ID,
            data = ., paste, collapse = "\n")


writexl::write_xlsx(df2,'C:/Users/thiago.oliveira/Downloads/AG 1-final.xlsx')
