rm(list = ls())
pacman::p_load('readxl','stringr', 'dplyr')

unirTexto = function(delimitador,coluna1, coluna2){
  linhas1 <- length(coluna1)
  linhas2 <- length(coluna2)
  colunaFinal <- c()
  i = 1
  if(linhas1 == linhas2){
    while(i <= linhas1){
      if(is.na(coluna1[i]) == TRUE){
        coluna1[i] = "celula vazia"
      }
      if(is.na(coluna2[i]) == TRUE){
        coluna2[i] = "celula vazia"
      }
      if(coluna1[i] != "celula vazia" &&
         coluna2[i] != "celula vazia"){
        coluna1[i] <- str_c(coluna1[i], delimitador, coluna2[i])
        
      }else if(coluna1[i] == "celula vazia" &&
               coluna2[i] != "celula vazia"){
        coluna1[i] = coluna2[i]
      }
      colunaFinal <- append(colunaFinal, coluna1[i])
      i= i + 1
      
    }
  } else {
    return("Quantidade de linhas diferente")
  }
  return(colunaFinal)
}

colunaObservacao = function(titulo, coluna){
  linhas <- length(coluna)
  i = 1
  colunaFinal <- c()
  while (i <= linhas) {
    if(is.na(coluna[i]) == FALSE){
      coluna[i] <- str_c(titulo, coluna[i])
    } else {
      coluna[i] <- "celula vazia"
    }
    if(str_detect(coluna[i], "celula vazia") == TRUE){
      coluna[i] = NA
    }
    colunaFinal <- append(colunaFinal, coluna[i])
    i=i+1
  }
  return(colunaFinal)
}

dir <- "C:/Users/thiago.oliveira/OneDrive/Área de Trabalho/Migração"
diretorio <- str_c(dir,"/input/")
diretorioFinal <- str_c(dir,"/output/")
arquivos <- list.files(diretorio, pattern = "\\.xlsx")

#tratando arquivos
for (arquivo in arquivos){
  if(str_detect(arquivo, "Appointments") == TRUE ){
    
    appointments <- str_c(diretorio, arquivo)
    
    df1 <- readxl::read_xlsx(appointments)
    
    dfFinal1 <- data.frame(df1$title, df1$start, df1$insuranceName, df1$comments,
                           df1$serviceName, df1$patientId, df1$scheduleId, df1$duration)
    
    dfFinal1$df1.start <- gsub("T", " ", dfFinal1$df1.start)
    
    dfFinal1$df1.start <- substr(dfFinal1$df1.start, 1,16)
    
    finalPath <- (str_c(diretorioFinal, arquivo))
    
    writexl::write_xlsx(dfFinal1,finalPath)
    
  }
  if(str_detect(arquivo, "Patients") == TRUE ){
    
    patients <- str_c(diretorio, arquivo)
    
    df2 <- readxl::read_xlsx(patients)
    
    patientAllergies <- (df2$patientAllergies) 
    df2Linhas <- length(patientAllergies)
    l = 1
    while (l<=df2Linhas) {
      if(patientAllergies[l] == "[]"){
        patientAllergies[l] = NA
      } else {
        textoJSON <- str_split(patientAllergies[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        patientAllergies[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    
    patientMedications <- (df2$patientMedications) 
    l = 1
    while (l<=df2Linhas) {
      if(patientMedications[l] == "[]"){
        patientMedications[l] = NA
      } else {
        textoJSON <- str_split(patientMedications[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        patientMedications[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    
    patientPrecedents <- (df2$patientPrecedents) 
    l = 1
    while (l<=df2Linhas) {
      if(patientPrecedents[l] == "[]"){
        patientPrecedents[l] = NA
      } else {
        textoJSON <- str_split(patientPrecedents[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        patientPrecedents[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    
    patientOtherMedicalInfo <- (df2$patientOtherMedicalInfo) 
    l = 1
    while (l<=df2Linhas) {
      if(patientOtherMedicalInfo[l] == "[]"){
        patientOtherMedicalInfo[l] = NA
      } else {
        textoJSON <- str_split(patientOtherMedicalInfo[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        patientOtherMedicalInfo[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    
    additionalDocuments <- (df2$additionalDocuments) 
    l = 1
    while (l<=df2Linhas) {
      if(additionalDocuments[l] == "[]"){
        additionalDocuments[l] = NA
      } else {
        textoJSON <- str_split(additionalDocuments[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        additionalDocuments[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    nomeCompleto <- unirTexto(" ", df2$firstName, df2$lastName)
    
    
    dfFinal2 <- data.frame(df2$id, nomeCompleto, df2$dateOfBirth, df2$phone, df2$alternatePhone,
                           df2$email, df2$street, df2$city, df2$postalCode, df2$gender,
                           df2$allergies, df2$history, df2$medication, df2$document, df2$observations, 
                           df2$neighborhood, df2$religion, df2$profession, df2$nationalHealthcareCardNumber,
                           df2$streetNumber, df2$country, df2$state, df2$referer, df2$bornInState, df2$bornInCity,
                           df2$nationality, df2$insurances.name, df2$insuranceCardNumber, patientAllergies,
                           patientMedications, patientPrecedents, patientOtherMedicalInfo, additionalDocuments)
    
    dfFinal2$df2.allergies <- colunaObservacao("Alergias: ", dfFinal2$df2.allergies)
    dfFinal2$df2.history <- colunaObservacao("Histórico: ", dfFinal2$df2.history)
    dfFinal2$df2.medication <- colunaObservacao("Medicação: ", dfFinal2$df2.medication)
    dfFinal2$df2.insuranceCardNumber <- colunaObservacao("Número do cartão de saúde: ", dfFinal2$df2.insuranceCardNumber)
    dfFinal2$df2.observations <- colunaObservacao("Observações: ", dfFinal2$df2.observations)
    dfFinal2$df2.nationalHealthcareCardNumber <- colunaObservacao("Número do cartão nacional de saúde: ", dfFinal2$df2.nationalHealthcareCardNumber)
    dfFinal2$df2.referer <- colunaObservacao("Indicação: ", dfFinal2$df2.referer)
    dfFinal2$df2.bornInState <- colunaObservacao("Estado que nasceu: ", dfFinal2$df2.bornInState)
    dfFinal2$df2.bornInCity <- colunaObservacao("Cidade que nasceu: ", dfFinal2$df2.bornInCity)
    dfFinal2$df2.nationality <- colunaObservacao("Nacionalidade: ", dfFinal2$df2.nationality)
    dfFinal2$df2.insurances.name <- colunaObservacao("Nome do seguro: ", dfFinal2$df2.insurances.name)
    dfFinal2$patientAllergies <- colunaObservacao("Alergias: ", dfFinal2$patientAllergies)
    dfFinal2$patientMedications <- colunaObservacao("Medicações: ", dfFinal2$patientMedications)
    dfFinal2$patientPrecedents <- colunaObservacao("Precedentes: ", dfFinal2$patientPrecedents)
    dfFinal2$patientOtherMedicalInfo <- colunaObservacao("Outras informações médicas: ", dfFinal2$patientOtherMedicalInfo)
    
    contador <- 1
    observacoes <- c()
    
    while(contador <= df2Linhas){
      observacoes[contador] <- NA
      if(is.na(dfFinal2$df2.allergies[contador]) == FALSE){
        observacoes[contador] <- dfFinal2$df2.allergies[contador]
      }
      
      # HISTORY ###########################
      if(is.na(dfFinal2$df2.history[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.history[contador]
        
      } else if(is.na(dfFinal2$df2.history[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.history[contador])
      }
      # Medication ###########################
      if(is.na(dfFinal2$df2.medication[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.medication[contador]
        
      } else if(is.na(dfFinal2$df2.medication[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.medication[contador])
      }
      
      # Insurance Card number #############
      if(is.na(dfFinal2$df2.insuranceCardNumber[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.insuranceCardNumber[contador]
        
      } else if(is.na(dfFinal2$df2.insuranceCardNumber[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.insuranceCardNumber[contador])
      }
      
      # Observations #############
      if(is.na(dfFinal2$df2.observations[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.observations[contador]
        
      } else if(is.na(dfFinal2$df2.observations[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.observations[contador])
      }
      
      # National health care Number #############
      if(is.na(dfFinal2$df2.nationalHealthcareCardNumber[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.nationalHealthcareCardNumber[contador]
        
      } else if(is.na(dfFinal2$df2.nationalHealthcareCardNumber[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.nationalHealthcareCardNumber[contador])
      }
      
      # Referer #############
      if(is.na(dfFinal2$df2.referer[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.referer[contador]
        
      } else if(is.na(dfFinal2$df2.referer[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.referer[contador])
      }
      
      # Born in state #############
      if(is.na(dfFinal2$df2.bornInState[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.bornInState[contador]
        
      } else if(is.na(dfFinal2$df2.bornInState[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.bornInState[contador])
      }
      
      # Born in city #############
      if(is.na(dfFinal2$df2.bornInCity[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.bornInCity[contador]
        
      } else if(is.na(dfFinal2$df2.bornInCity[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.bornInCity[contador])
      }
      
      # Nationality #############
      if(is.na(dfFinal2$df2.nationality[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$df2.nationality[contador]
        
      } else if(is.na(dfFinal2$df2.nationality[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.nationality[contador])
      }
      
      # Insurances name #############
      if(is.na(dfFinal2$df2.insurances.name[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE &&
         dfFinal2$df2.insurances.name[contador] != "Nome do seguro: sin-aseguradora"){
        
        observacoes[contador] <- dfFinal2$df2.insurances.name[contador]
        
      } else if(is.na(dfFinal2$df2.insurances.name[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE &&
                dfFinal2$df2.insurances.name[contador] != "Nome do seguro: sin-aseguradora"){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$df2.insurances.name[contador])
      }
      
      # Patient Allergies #############
      if(is.na(dfFinal2$patientAllergies [contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$patientAllergies[contador]
        
      } else if(is.na(dfFinal2$patientAllergies[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$patientAllergies[contador])
      }
      
      # Patient Medications #############
      if(is.na(dfFinal2$patientMedications[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$patientMedications[contador]
        
      } else if(is.na(dfFinal2$patientMedications[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$patientMedications[contador])
      }
      
      # Patient Precedents #############
      if(is.na(dfFinal2$patientPrecedents[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$patientPrecedents[contador]
        
      } else if(is.na(dfFinal2$patientPrecedents[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$patientPrecedents[contador])
      }
      
      # Patient Other Medical Info #############
      if(is.na(dfFinal2$patientOtherMedicalInfo[contador]) == FALSE &&
         is.na(observacoes[contador]) == TRUE){
        
        observacoes[contador] <- dfFinal2$patientOtherMedicalInfo[contador]
        
      } else if(is.na(dfFinal2$patientOtherMedicalInfo[contador]) == FALSE &&
                is.na(observacoes[contador]) == FALSE){
        observacoes[contador] <- str_c(observacoes[contador], "\n", dfFinal2$patientOtherMedicalInfo[contador])
      }
      
      contador <- contador + 1
    }
    
    dfFinal2 <- data.frame(df2$id, nomeCompleto, df2$dateOfBirth, df2$phone, df2$alternatePhone,
                           df2$email, df2$street, df2$city, df2$postalCode, df2$gender,
                           df2$document,df2$neighborhood, df2$religion, df2$profession,
                           df2$streetNumber, df2$country, df2$state, df2$nationality,
                           additionalDocuments, observacoes)
    
    
    
    
    dfFinal2$df2.dateOfBirth <- gsub("T", " ", dfFinal2$df2.dateOfBirth)
      
    dfFinal2$df2.dateOfBirth <- gsub("T", " ", dfFinal2$df2.dateOfBirth)
    
    dfFinal2$df2.dateOfBirth <- substr(dfFinal2$df2.dateOfBirth, 1,10)
    
    dfFinal2 %>% 
      mutate(df2.dateOfBirth = as.Date(df2.dateOfBirth))
    
    
    dfFinal2$df2.phone <- gsub("+55", "", dfFinal2$df2.phone)
    
    dfFinal2$df2.alternatePhone <- gsub("+55", "", dfFinal2$df2.alternatePhone)
    
    finalPath <- (str_c(diretorioFinal, arquivo))
    
    writexl::write_xlsx(dfFinal2,finalPath )
    
  }
  if(str_detect(arquivo, "EHR") == TRUE ){
    
    ehr <- str_c(diretorio, arquivo)
    
    df3 <- readxl::read_xlsx(ehr)
    
    dfFinal3 <- data.frame(df3$PatientId, df3$date, df3$customFields)
    
    customFields <- dfFinal3$df3.customFields
    
    itens <- length(customFields)
    x = 1
    
    while(x <= itens){
      
      item <- customFields[x]
      customFields[x] <- NA
      
      ### Values
      stringValues <- str_split(item,"'value': ")
      stringLabels <- str_split(item,"label': '"  )
      itensValue <- length(stringValues[[1]])
      y = 2
      while(y<=itensValue){
        stringValue <- str_split(stringValues[[1]][y], ", 'sequenceNumber'")
        stringLabel <- str_split(stringLabels[[1]][y], "', 'type':")
        value <- stringValue[[1]][1]
        label <- stringLabel[[1]][1]
        if(is.na(customFields[x]) == TRUE && value != "None"){
          customFields[x] <- label
          customFields[x] <- str_c(customFields[x], "\n", value)
          label1 <- TRUE
        }else if(value!= "None"){
          customFields[x] <- str_c(customFields[x], "\n", label)
          customFields[x] <- str_c(customFields[x], "\n", value)
        }
        y = y + 1
      }
      x = x + 1
    }
    
    dfFinal3$df3.customFields <- customFields
    
    dfFinal3$df3.customFields <- gsub('"value": "', "", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub(',"codes":', " ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub("]}'", " ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub('"codes":'," ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub(']'," ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub('\\['," ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub('\\{'," ", dfFinal3$df3.customFields)
    dfFinal3$df3.customFields <- gsub('\\}'," ", dfFinal3$df3.customFields)
    
    finalPath <- (str_c(diretorioFinal, arquivo))
    
    writexl::write_xlsx(dfFinal3,finalPath)
  }
}

# Apagando colunas vazias

for (arquivo in arquivos){
  arquivo <- str_c(diretorioFinal, arquivo)
  df <-readxl::read_xlsx(arquivo)
  dfLines <- nrow(df)
  dfCols = ncol(df)
  c = 1
  lista <- c();
  if(dfLines == 0) {
    file.remove(arquivo)
    print("Arquivo vazio")
  }else {
    while(c <= dfCols){
      l = 1
      hasValue = FALSE;
      while(l<=dfLines){
        if(is.na(df[l,c]) == FALSE){
          lista <- append(lista, c)
          break;
        }
        l = l+1
      }
      c = c+1
    }
    c = dfCols
    while(c > 0) {
      if(c %in% lista == FALSE){
        df[c] = NULL;
      }
      c = c - 1
    }
    writexl::write_xlsx(df,arquivo)
  }
}
