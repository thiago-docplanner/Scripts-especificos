rm(list = ls())

pacman::p_load('dplyr','purrr','readxl','stringr')

df <-readxl::read_xlsx('C:/Users/thiago.oliveira/Downloads/EHR CAMILA.xlsx')
df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)

# criando os valores nos espacos nulos das colunas ate que o proximo valor apareca

nome_instituicao <- c()
y <- 1
for (i in df$nome_instituicao) {
  if (is.na(i) == FALSE) {
    nome_instituicao[y] <- i
    x <- i
    y = y+ 1
  } else {
    nome_instituicao[y] <- x
    y = y+ 1
  }
}

nome_profissional <- c()
y <- 1
for (i in df$nome_profissional) {
  if (is.na(i) == FALSE) {
    nome_profissional[y] <- i
    x <- i
    y = y+ 1
  } else {
    nome_profissional[y] <- x
    y = y+ 1
  }
}

dia_emissao <- c()
y <- 1
for (i in df$dia_emissao) {
  if (is.na(i) == FALSE) {
    dia_emissao[y] <- i
    x <- i
    y = y+ 1
  } else {
    dia_emissao[y] <- x
    y = y+ 1
  }
}

cdoc_tx_nome <- c()
y <- 1
for (i in df$cdoc_tx_nome) {
  if (is.na(i) == FALSE) {
    cdoc_tx_nome[y] <- i
    x <- i
    y = y+ 1
  } else {
    cdoc_tx_nome[y] <- x
    y = y+ 1
  }
}

pess_id <- c()
y <- 1
for (i in df$pess_id) {
  if (is.na(i) == FALSE) {
    pess_id[y] <- i
    x <- i
    y = y+ 1
  } else {
    pess_id[y] <- x
    y = y+ 1
  }
}

pess_tx_nome <- c()
y <- 1
for (i in df$pess_tx_nome) {
  if (is.na(i) == FALSE) {
    pess_tx_nome[y] <- i
    x <- i
    y = y+ 1
  } else {
    pess_tx_nome[y] <- x
    y = y+ 1
  }
}

titulo_documento <- c()
y <- 1
for (i in df$titulo_documento) {
  if (is.na(i) == FALSE) {
    titulo_documento[y] <- i
    x <- i
    y = y+ 1
  } else {
    titulo_documento[y] <- x
    y = y+ 1
  }
}



# Criando uma nova tabela com as colunas criadas e as demais necessarias

df2 <- data.frame(nome_instituicao,nome_profissional, dia_emissao, cdoc_tx_nome,
                  pess_id, pess_tx_nome, titulo_documento,df$conteudo)

# Mudando o nome das colunas que j? existiam

names(df2)[names(df2) == 'df.conteudo'] <- 'conteudo'


#### Tratamento ####

df3 <- df2 %>%
  aggregate(conteudo ~ nome_instituicao + nome_profissional +  dia_emissao +  cdoc_tx_nome +
            pess_id + pess_tx_nome + titulo_documento,
            data = ., paste, collapse = "\n")


writexl::write_xlsx(df3,'C:/Users/thiago.oliveira/Downloads/EHR CAMILA-final.xlsx')
