# Limpando o ambiente, importando bibliotecas e tabelas e criando DF final

rm(list = ls())

pacman::p_load('dplyr','purrr','readxl','stringr')

df <-readxl::read_xlsx('C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/FormatterBR/pacientes.xlsx')

df2 <-data.frame()

df_final <- data.frame()%>%
  mutate("NOME",	"DT. CADASTRO",	"PRONTUARIO",	"DA. NASCIMENTO",	"COR",	
         "CONVENIO",	"ENDEREÇO",	"ESTADO CIVIL",	"PLANO/EMPRESA",
         "BAIRRO",	"PROFISSAO",	"NUMERO DA MATRICULA / CARTEIRA",	"CIDADE",
         "IDENTIDADE",	"DATA DE VALIDADE",	"CEP",	"UF",	
         "CPF",	"TITULAR","CELULAR",	"WHATSAPP",	"RESIDENCIAL","COMERCIAL",
         "EMAIL", "NUMERO DO CARTAO NACIONAL DE SAUDE")

# Deixando todos NA com o valor "-----"
colunas <- ncol(df)
linhas <- nrow(df)
l = 1

for (y in 1:linhas){
  i <- 1
  c <- 1
  while (i < colunas+1){
    if (is.na(df[y,i]) == FALSE) {
      df2[l, c] <- df[y,i]
    } else {
      df2[l, c] <- "-----"
    }
    i = (i + 1)
    c = (c + 1)
  }
  l = (l + 1)
}

# Jogando valores na Df_final
colunas <- ncol(df2)
linhas <- nrow(df2)
l = 0

for (y in 1:linhas){
  i <- 1
  while (i < colunas+1){
    if (df2[y,i] == "NOME:"){
      l=l+1 
      c <- 1
    }else if (df2[y,i] == "SEXO:" || df2[y,i] ==	"DATA DE CADASTRO:"
              || df2[y,i] == "PRONTUÁRIO:,DT."	|| df2[y,i] == "DT. NASCIMENTO:"	|| df2[y,i] == "COR:"
              || df2[y,i] == "CONVÊNIO:" || df2[y,i] ==	"ENDEREÇO:" || df2[y,i] ==	"ESTADO CIVIL:"
              || df2[y,i] == "PLANO / EMPRESA:" || df2[y,i] == "BAIRRO:"
              || df2[y,i] == "PROFISSÃO:" || df2[y,i] ==	"Nº MAT. / CARTEIRA:" || df2[y,i] ==	"CIDADE:"
              || df2[y,i] == "IDENTIDADE:" || df2[y,i] ==	"DATA VALIDADE:" || df2[y,i] ==	"CEP:"
              || df2[y,i] == "UF:" || df2[y,i] ==	"CPF:" 
              || df2[y,i] ==	"TITULAR:"  || df2[y,i] == "Celular :" || df2[y,i] == "WhatsApp :"
              || df2[y,i] ==	"Residencial :"  || df2[y,i] == "Comercial :" || df2[y,i] == "E-MAIL:"
              || df2[y,i] ==	"Nº CARTÃO NAC. DE SAÚDE:"){
      
      c <- c +1
    } else if ( df2[y,i] != "-----" ){
      df_final[l,c] <- df2[y,i]
    }
    i = (i + 1)
  }
}

writexl::write_xlsx(df_final,'C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/FormatterBR/pacientes-final.xlsx')
