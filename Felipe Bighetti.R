rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr")

diretorio <- "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Felipe Bighetti/documentos"
setwd(diretorio)
dfFinal<- data.frame() %>%
  mutate("Data Atendimento", "Número de prontuário", "CPF", "Nº Cartão",
         "Paciente", "Nome da mãe", "Sexo", "Data de nascimento", "Telefone",
         "Celular","Cidade", "Logradouro", "Bairro", "Complemento", "CEP",
         "Classificação", "Profissional de saúde", "Local de atendimento", "Tipo de atendimento",
         "Anamnese")


y = 0

arquivos <- list.files(diretorio, pattern = "\\.xlsx")

for (arquivo in arquivos){
  
  arquivo <- str_c(diretorio, "/", arquivo)
  
  df <-readxl::read_xlsx(arquivo)
  df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)
  
  linhas = nrow(df)
  colunas = ncol(df)
  l = 1
  while(l<=linhas){
    c = 1
    while(c<=colunas){
      if(is.na(df[l,c]) == TRUE){
        df[l,c] = "celula vazia"
      }
      if(str_detect(df[l,c], "Data Atendimento") == TRUE ){
        x = 1
        y = y + 1
        anamnese = ""
      }
      
      else if(str_detect(df[l,c], "CPF:") == TRUE){
        x = 3
      }
      
      else if(str_detect(df[l,c], "N.Cartão:") == TRUE){
        x = 4
      }
      
      else if(str_detect(df[l,c], "Paciente:") == TRUE){
        x = 5
      }
      
      else if(str_detect(df[l,c], "Nome Mãe:") == TRUE){
        x = 6
      }
      
      else if(str_detect(df[l,c], "Sexo:") == TRUE){
        x = 7
      }
      
      else if(str_detect(df[l,c], "Data Nascimento:") == TRUE){
        x = 8
      }
      
      else if(str_detect(df[l,c], "Telefone:") == TRUE){
        x = 9
      }
      
      else if(str_detect(df[l,c], "Celular:") == TRUE){
        x = 10
      }
      
      else if(str_detect(df[l,c], "Cidade:") == TRUE){
        x = 11
      }
      
      else if(str_detect(df[l,c], "Logradouro:") == TRUE){
        x = 12
      }
      
      else if(str_detect(df[l,c], "Bairro:") == TRUE){
        x = 13
      }
      
      else if(str_detect(df[l,c], "Complemento:") == TRUE){
        x = 14
      }
      
      else if(str_detect(df[l,c], "CEP:") == TRUE){
        x = 15
      }
      
      else if(str_detect(df[l,c], "Classificação:") == TRUE){
        x = 16
        dfFinal[y,x] = df[l,c]
      }
      
      else if(str_detect(df[l,c], "Profissional de saúde:") == TRUE){
        x = 17
      }
      
      else if(str_detect(df[l,c], "Local de atendimento:") == TRUE){
        x = 18
      }
      
      else if(str_detect(df[l,c], "Tipo de atendimento:") == TRUE){
        x = 19
      }
      
      else if(str_detect(df[l,c], "ANAMNESE") == TRUE){
        x = 20
      }
      
      else if(str_detect(df[l,c], "RES - ") == TRUE ||
              str_detect(df[l,c], "Pag.:") == TRUE ||
              str_detect(df[l,c], "Column") == TRUE){
        df[l,c] = "celula vazia"
      }
      
      else if(str_detect(df[l,c], "Data Atendimento") == FALSE &&
              str_detect(df[l,c], "Prontuário N") == FALSE &&
              str_detect(df[l,c], "CPF:") == FALSE &&
              str_detect(df[l,c], "N.Cartão:") == FALSE &&
              str_detect(df[l,c], "Paciente:") == FALSE &&
              str_detect(df[l,c], "Nome Mãe:") == FALSE &&
              str_detect(df[l,c], "Sexo:") == FALSE &&
              str_detect(df[l,c], "Data Nascimento:") == FALSE &&
              str_detect(df[l,c], "Telefone:") == FALSE &&
              str_detect(df[l,c], "Celular:") == FALSE &&
              str_detect(df[l,c], "Cidade:") == FALSE &&
              str_detect(df[l,c], "Logradouro:") == FALSE &&
              str_detect(df[l,c], "Bairro:") == FALSE &&
              str_detect(df[l,c], "Complemento:") == FALSE &&
              str_detect(df[l,c], "CEP:") == FALSE &&
              str_detect(df[l,c], "Classificação:") == FALSE &&
              str_detect(df[l,c], "Profissional de saúde:") == FALSE &&
              str_detect(df[l,c], "Local de atendimento:") == FALSE &&
              str_detect(df[l,c], "Tipo de atendimento:") == FALSE &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Código") == FALSE &&
              str_detect(df[l,c], "Descrição") == FALSE &&
              str_detect(df[l,c], "N.Solic") == FALSE &&
              str_detect(df[l,c], "Futura") == FALSE &&
              str_detect(df[l,c], "celula vazia") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE &&
              x != 20 &&
              x != 21) {
        
        dfFinal[y,x] = df[l,c]
        x = x + 1
      } 
      
      else if(x == 20 && 
              str_detect(df[l,c], "DIAGNOSTICO :") == TRUE){
        dfFinal[y,x] = "Diagnóstico: "
        anamnese = "Diagnostico"
      } 
      
      else if(x == 20 && 
              anamnese == "Diagnostico" &&
              df[l,c] != "celula vazia" &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Código") == FALSE &&
              str_detect(df[l,c], "Descrição") == FALSE &&
              str_detect(df[l,c], "N.Solic") == FALSE &&
              str_detect(df[l,c], "Futura") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE){
        
        dfFinal[y,x] = str_c(dfFinal[y,x], df[l,c])
        
      }
      
      else if(x == 20 && 
              str_detect(df[l,c], "CONDUTA") == TRUE){
        dfFinal[y,x] = str_c(dfFinal[y,x] , "\n\n" , "Conduta: ")
        anamnese = "Conduta"
      } 
      
      else if(x == 20 && 
              anamnese == "Conduta" &&
              df[l,c] != "celula vazia" &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Código") == FALSE &&
              str_detect(df[l,c], "Descrição") == FALSE &&
              str_detect(df[l,c], "N.Solic") == FALSE &&
              str_detect(df[l,c], "Futura") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE){
        
        dfFinal[y,x] = str_c(dfFinal[y,x], df[l,c])
        
      }
      
      else if(x == 20 && 
              str_detect(df[l,c], "EVOLUÇÃO") == TRUE){
        dfFinal[y,x] = str_c(dfFinal[y,x] , "\n\n" , "Evolução: ")
        anamnese = "Evolucão"
      } 
      
      else if(x == 20 && 
              anamnese == "Evolucão" &&
              df[l,c] != "celula vazia" &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Código") == FALSE &&
              str_detect(df[l,c], "Descrição") == FALSE &&
              str_detect(df[l,c], "N.Solic") == FALSE &&
              str_detect(df[l,c], "Futura") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE){
        
        dfFinal[y,x] = str_c(dfFinal[y,x], " ", df[l,c])
      }
      
      else if(x == 20 && 
              str_detect(df[l,c], "EXAMES") == TRUE){
        dfFinal[y,x] = str_c(dfFinal[y,x] , "\n\n" , "Exames: ")
        anamnese = "Exames"
      } 
      ### FALTA EXAMES ###
      else if(x == 20 && 
              anamnese == "Exames" &&
              df[l,c] != "celula vazia" &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE){
        if(str_detect(df[l,c], "Código") == TRUE){
          anamnese = "Codigo"
          quantidadeExames = 1
        } 
      }
      else if(x == 20 && 
              anamnese == "Codigo" &&
              str_detect(df[l,c], "ANAMNESE") == FALSE &&
              str_detect(df[l,c], "DIAGNOSTICO") == FALSE &&
              str_detect(df[l,c], "CONDUTA") == FALSE &&
              str_detect(df[l,c], "EVOLUÇÃO") == FALSE &&
              str_detect(df[l,c], "EXAMES") == FALSE &&
              str_detect(df[l,c], "Pag.:") == FALSE &&
              str_detect(df[l,c], "Column") == FALSE){
        
        if(c == 1 && 
           df[l,c] != "celula vazia" &&
           str_detect(df[l,c], "RES -") == FALSE){
          dfFinal[y,x] = str_c(dfFinal[y,x], "\n","Código: ", df[l,c])
          observacao = 0
          c = 2
          while(c <= colunas){
            if(is.na(df[l,c]) == TRUE){
              df[l,c] = "celula vazia"
            }
            if(df[l,c] != "celula vazia" && observacao == 0){
              descricao = df[l,c]
              observacao = 1
              
            } else if(df[l,c] != "celula vazia" && observacao == 1){
              dfFinal[y,x] = str_c(dfFinal[y,x], "\n","Número de solicitação: ", df[l,c])
              observacao = 2
              
            }else if(df[l,c] != "celula vazia" && observacao == 2){
              dfFinal[y,x] = str_c(dfFinal[y,x], "\n","Futura: ", df[l,c])
              observacao = 3
            }
            if(observacao == 3){
              dfFinal[y,x] = str_c(dfFinal[y,x], "\n","Descrição:", "\n", descricao)
              observacao = 0
            }
            
            c = c +1
          }
          c = 1
        } 
        else if(c == 1 &&
                df[l,c] == "celula vazia"){
          c = 2
          while(c <= colunas){
            if(is.na(df[l,c]) == TRUE){
              df[l,c] = "celula vazia"
            }
            
            if(df[l,c] != "celula vazia"){
              dfFinal[y,x] = str_c(dfFinal[y,x]," ", df[l,c])
            }
            c = c + 1
          }
          c = 1
        }
        
      }
      
      
      c = c +1
    }
    l = l +1
  }
}

writexl::write_xlsx(dfFinal, "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Felipe Bighetti/teste.xlsx")
