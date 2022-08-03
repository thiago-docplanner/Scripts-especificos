rm(list = ls())
pacman::p_load('readxl','stringr', 'dplyr')

dir <- getwd()
diretorio <- str_c(dir,"/input/")
diretorioFinal <- str_c(dir,"/output/")
arquivos <- list.files(diretorio, pattern = "\\.xlsx")

for (arquivo in arquivos){
  if(str_detect(arquivo, "Appointments") == TRUE ){
    
    appointments <- str_c(diretorio, arquivo)
    
    df1 <- readxl::read_xlsx(appointments)
    
    dfFinal1 <- data.frame(df1$title, df1$start, df1$end, df1$insuranceName, df1$comments,
                           df1$serviceName, df1$patientId, df1$scheduleId, df1$duration)
    
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
      print(l)
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
      print(l)
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
      print(l)
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
      print(l)
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
      print(l)
      if(additionalDocuments[l] == "[]"){
        additionalDocuments[l] = NA
      } else {
        textoJSON <- str_split(additionalDocuments[l], "'description': '" )
        textoFinal <- str_split(textoJSON[[1]][2], "'\\}")
        additionalDocuments[l] <- textoFinal[[1]][1]
        
      }
      l = l + 1
    }
    
    dfFinal2 <- data.frame(df2$id, df2$firstName, df2$lastName, df2$dateOfBirth, df2$phone, df2$email,
                           df2$email, df2$alternatePhone, df2$street, df2$city, df2$postalCode, df2$gender,
                           df2$allergies, df2$history, df2$medication, df2$document, df2$observations, 
                           df2$neighborhood, df2$religion, df2$profession, df2$nationalHealthcareCardNumber,
                           df2$streetNumber, df2$country, df2$state, df2$referer, df2$bornInState, df2$bornInCity,
                           df2$nationality, df2$insurances.name, df2$insuranceCardNumber, patientAllergies,
                           patientMedications, patientPrecedents, patientOtherMedicalInfo, additionalDocuments)
    
    finalPath <- (str_c(diretorioFinal, arquivo))
    
    writexl::write_xlsx(dfFinal2,finalPath )
    
  }
  if(str_detect(arquivo, "EHR") == TRUE ){
    
    ehr <- str_c(diretorio, arquivo)
    
    df3 <- readxl::read_xlsx(ehr)
    
    dfFinal3 <- data.frame(df3$PatientId, df3$EHRs, df3$date, df3$customFields)
    
    finalPath <- (str_c(diretorioFinal, arquivo))
    
    writexl::write_xlsx(dfFinal3,finalPath )
  }
}

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
