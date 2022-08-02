rm(list = ls())

pacman::p_load('jsonlite','dplyr','purrr','readxl','stringr','tibble')

setwd('C:/Users/Doctoralia/Desktop/Doctoralia/Textos e scripts/')

presp1 = presp3 = block =  attest = exam =  imagem = 
  tb_insert_attest = tb_insert_exam = tb_insert_presp1 = tb_insert_presp3 = 
  tbinsert = pacientes = tb_exam = tb_attest = tb_block = tb_presp1 = tb_presp3 =
  Attest_Desc = Block_Desc = Presp1_Desc = Presp3_Desc = Exam_Desc = NULL

df <-read.csv('Andrea 26-07-2022-event_record.csv', sep=',', encoding = 'UTF-8') %>% 
  select(-physician_name,-extra_pack,-procedure_pack,-healthinsurance_pack) %>% 
  mutate(eventblock_pack = gsub("json::", "", eventblock_pack)) %>% 
  mutate(eventblock_pack = as.character(eventblock_pack)) %>% 
  mutate( P3 = case_when(
    str_detect(eventblock_pack, '"prescription_v3"')  ~ 'Prescription3',
    TRUE ~  "Outros"),
    P1 = case_when(
      str_detect(eventblock_pack, '"prescription_v1"')  ~ 'Prescription1',
      TRUE ~  "Outros"),
    EX = case_when(
      str_detect(eventblock_pack, '"exam_request"')  ~ 'Exam',
      TRUE ~  "Outros"),
    AT = case_when(
      str_detect(eventblock_pack, '"attest"')  ~ 'Attest',
      TRUE ~  "Outros"),
    BL = case_when(
      str_detect(eventblock_pack, '"block"')  ~ 'Block',
      TRUE ~  "Outros"),
    IM = case_when(
      str_detect(eventblock_pack, '"image"')  ~ 'Imagem', 
      TRUE ~  "Outros")
    )

dadospacientes <- df %>% 
  select(pk,patient_id,patient_name,date,start_time,end_time)

#### Verificando se as colunas existem na tabela ####

existeP1 = FALSE

for(numero in df$P1){
  if (numero != "Outros"){
    existeP1 = TRUE
  }
}

existeP3 = FALSE

for(numero in df$P3){
  if (numero != "Outros"){
    existeP3 = TRUE
  }
}

existeEX = FALSE

for(numero in df$EX){
  if (numero != "Outros"){
    existeEX = TRUE
  }
}

existeAT = FALSE

for(numero in df$AT){
  if (numero != "Outros"){
    existeAT = TRUE
  }
}

existeBL = FALSE

for(numero in df$BL){
  if (numero != "Outros"){
    existeBL = TRUE
  }
}

existeIM = FALSE

for(numero in df$IM){
  if (numero != "Outros"){
    existeIM = TRUE
  }
}




#### TIPOS ####

pacientes <- df %>% 
  select(pk,patient_id)

if (existeP1 == TRUE){
  presp1 <- df %>% 
    filter(.,P1=="Prescription1")
}


if (existeP3 == TRUE){
  presp3 <- df %>% 
    filter(.,P3=="Prescription3")
}

if (existeEX == TRUE){
  exam <- df %>% 
    filter(.,EX=="Exam")
}

if (existeAT == TRUE){
  attest <- df %>% 
    filter(.,AT=="Attest")
}


if (existeBL == TRUE){
  block <- df %>% 
    filter(.,BL=="Block")
}

if (existeIM == TRUE){
  imagem <- df %>% 
    filter(.,IM=="Imagem")
}

#outros <- df %>% 
 # filter(.,FLAG=="Outros")


#### CRIANDO DF ####
#### imagem ####

if (existeIM == TRUE){
  tb_imagem <- tibble()

  for (i in 1:nrow(block)){
    x <- imagem$eventblock_pack[i]
    tbinsert_imagem <-fromJSON(x, simplifyVector = FALSE) %>% 
      pluck('image') %>% 
      map_dfr(as.list) %>% 
      mutate(pk = imagem$pk[i],
            patient_name = imagem$patient_name[i],
            start_time = imagem$start_time[i],
            end_time = imagem$end_time[i],
           datadia = imagem$date[i]) %>% 
      select(pk,patient_name,start_time,end_time,datadia,image) %>% 
      mutate(image = 'Contem imagem') 
    tb_imagem <- tb_imagem %>% 
      bind_rows(tbinsert_imagem)  
  }

  tb_imagem <- tb_imagem %>% 
    unique() %>% 
    left_join(pacientes,by="pk") 
}

#### block ####
if (existeBL == TRUE){
  tb_block <- tibble()

  for (i in 1:nrow(block)){
    x <- block$eventblock_pack[i]
    tbinsert <-fromJSON(x, simplifyVector = TRUE) %>% 
      pluck('block') %>% 
      select(-date_added,-kind,-ordering) %>% 
      rename(., nome = name) %>%
      rename(., descricao = value) %>% 
      mutate(pk = block$pk[i],
            patient_name = block$patient_name[i],
            start_time = block$start_time[i],
            end_time = block$end_time[i],
            datadia = block$date[i],
            descricao = str_remove_all(.$descricao,'(<p>|</p>|<br/>|<br>)')
            ) %>%
      rename(., paciente = patient_name, hora_inicio = start_time, hora_fim = end_time) %>% 
      mutate(diagnostico = paste(nome,descricao,'\n')) %>% 
      select(-nome,-tab,-descricao) %>% 
      aggregate(diagnostico ~ pk + paciente + hora_inicio + hora_fim + datadia,
                data = ., paste, collapse = "\n")
    tb_block <- tb_block %>% 
      bind_rows(tbinsert)  
    }

  tb_block <- tb_block %>% 
    left_join(pacientes,by="pk") 
}

#### attest ####
if (existeAT == TRUE){
  tb_attest <- tibble()

  for (i in 1:nrow(attest)){
    x <- attest$eventblock_pack[i]
    tb_insert_attest <-fromJSON(x, simplifyVector = TRUE) %>% 
      pluck('attest') %>% 
      select(-date_added,-print_date) %>% 
      rename(., nome = name,descricao = value) %>% 
      mutate(pk = attest$pk[i],
            patient_name = attest$patient_name[i],
            start_time = attest$start_time[i],
            end_time = attest$end_time[i],
            datadia = attest$date[i]) %>% 
      rename(., paciente = patient_name, hora_inicio = start_time, hora_fim = end_time) %>% 
      mutate(diagnostico = paste(nome,descricao,'\n')) %>% 
      select(-nome,-tab,-descricao) %>% 
      aggregate(diagnostico ~ pk + paciente + hora_inicio + hora_fim + datadia,
                data = ., paste, collapse = "\n") 
    tb_attest <- tb_attest %>% 
      bind_rows(tb_insert_attest) %>% 
      mutate(diagnostico = str_remove_all(.$diagnostico,'(<p style="text-align: center;">|<p>|</p>|<br/>|<br>|<b>|</b>|<p style="text-align: justify;">|&nbsp;|[A-Za-z0-9]{20,})'))
    }

  tb_attest <- tb_attest %>% 
    left_join(pacientes,by="pk") 
}


#### exam ####

if (existeEX == TRUE){
  tb_exam <- tibble()

  for (i in 1:nrow(exam)){
    x <- exam$eventblock_pack[i]
    tb_insert_exam <-fromJSON(x, simplifyVector = TRUE) %>% 
      pluck('exam_request') %>% 
      mutate(., clinical_indication = if (exists('clinical_indication', where = .)) clinical_indication else NA,
            date_added = if (exists('date_added', where = .)) date_added else NA,
            exam_type = if (exists('exam_type', where = .)) exam_type else NA,
            is_model = if (exists('is_model', where = .)) is_model else NA,
            items = if (exists('items', where = .)) items else NA,
            name = if (exists('name', where = .)) name else NA,
            ordering = if (exists('ordering', where = .)) ordering else NA,
            physician_id = if (exists('physician_id', where = .)) physician_id else NA,
            print_date = if (exists('print_date', where = .)) print_date else NA,
            tab = if (exists('tab', where = .)) tab else NA) %>% 
      select(-date_added,-ordering,-print_date,-physician_id,-exam_type,-is_model) %>% 
      rename(., nome = name) %>% 
      mutate(items = gsub("list", "", items),
            tipo_exame = str_extract(items,pattern = '(?<= text )(.*)'),
            pk = exam$pk[i],
            patient_name = exam$patient_name[i],
            start_time = exam$start_time[i],
            end_time = exam$end_time[i],
            datadia = exam$date[i],
            ) %>% 
      select(-items) %>% 
      rename(., paciente = patient_name, hora_inicio = start_time, hora_fim = end_time) %>% 
      mutate(descricao = paste(nome,tipo_exame,'\n')) %>%
      select(-nome,-tab,-clinical_indication,-tipo_exame) %>% 
      aggregate(descricao ~ pk + paciente + hora_inicio + hora_fim + datadia,
                data = ., paste, collapse = "\n")
    tb_exam <- tb_exam %>% 
      bind_rows(tb_insert_exam) %>% 
      mutate(descricao = str_remove_all(.$descricao,'[0-9]{8}|("\\))'))
    }

  tb_exam <- tb_exam %>% 
    left_join(pacientes,by="pk")
}


#### presp3 ####
if (existeP3 == TRUE){
  tb_presp3 <- tibble()

  for (i in 1:nrow(presp3)){
    x <- presp3$eventblock_pack[i]
    tb_insert_presp3 <-fromJSON(x, simplifyVector = FALSE) %>% 
      pluck('prescription_v3') %>% 
      map('items') %>% 
      map('drugs') %>% 
      map_dfr(as.list) %>% 
      select(-controlled,-drugKind,-id,-integration) %>% 
      mutate(., quantity = if (exists('quantity', where = .)) quantity else NA,
            posology = if (exists('posology', where = .)) posology else NA,
            description = if (exists('description', where = .)) description else NA,
            name = if (exists('name', where = .)) name else NA) %>% 
      rename(., descricao = description,
            nome = name,
            posologia = posology,
            quantidade = quantity) %>% 
      mutate(posologia = str_remove_all(.$posologia,'(<p>|</p>|<br>)'),
            pk = presp3$pk[i],
            patient_name = presp3$patient_name[i],
            start_time = presp3$start_time[i],
            end_time = presp3$end_time[i],
            datadia = presp3$date[i]
            ) %>% 
      rename(., paciente = patient_name, hora_inicio = start_time, hora_fim = end_time) %>% 
      mutate(receita = paste('Medicação:',descricao,
                            "\n",
                            'Nome:',nome,
                            "\n",
                            'Posologia:',posologia,
                            "\n",
                            'Quantidade:',quantidade, 
                            sep="\n")) %>%
      select(-descricao,-nome,-posologia,-quantidade) %>% 
      aggregate(receita ~ pk + paciente + hora_inicio + hora_fim + datadia,
                data = ., paste, collapse = "\n")
    tb_presp3 <- tb_presp3 %>% 
      bind_rows(tb_insert_presp3) %>%
      select(pk,paciente,hora_inicio,hora_fim,datadia,receita) 
  }

  tb_presp3 <- tb_presp3 %>% 
    left_join(pacientes,by="pk") 
}

#### presp1 ####
if (existeP1 == TRUE){
  tb_presp1 <- tibble()

  for (i in 1:nrow(presp1)){
    x <- presp1$eventblock_pack[i]
    tb_insert_presp1 <-fromJSON(x, simplifyVector = FALSE) %>% 
      pluck('prescription_v1') %>% 
      map('items') %>% 
      #map('drugs') %>% 
      map_dfr(as.list) %>% 
      select(-ordering,-kind) %>% 
      mutate(., quantity = if (exists('quantity', where = .)) quantity else NA,
            posology = if (exists('posology', where = .)) posology else NA,
            text = if (exists('text', where = .)) text else NA) %>% 
      rename(., texto = text,
            posologia = posology,
            quantidade = quantity) %>% 
      mutate(posologia = str_remove_all(.$posologia,'(<p>|</p>|<br>)'),
            pk = presp1$pk[i],
            patient_name = presp1$patient_name[i],
            start_time = presp1$start_time[i],
            end_time = presp1$end_time[i],
            datadia = presp1$date[i]
      ) %>% 
      rename(., paciente = patient_name, hora_inicio = start_time, hora_fim = end_time) %>% 
      mutate(receita = paste('Medicação:',texto,
                            "\n",
                            'Posologia:',posologia,
                            "\n",
                            'Quantidade:',quantidade, 
                            sep="\n")) %>%
      select(-texto,-posologia,-quantidade) %>% 
      aggregate(receita ~ pk + paciente + hora_inicio + hora_fim + datadia,
                data = ., paste, collapse = "\n")
    tb_presp1 <- tb_presp1 %>% 
      bind_rows(tb_insert_presp1) %>% 
      select(pk,paciente,hora_inicio,hora_fim,datadia,receita) 
  }

  tb_presp1 <- tb_presp1 %>% 
    left_join(pacientes,by="pk")
}

### edicao final ####

rm(tb_insert_attest,tb_insert_exam,tb_insert_presp1,tb_insert_presp3,tbinsert,
   presp3,exam,pacientes,attest)

lista1 <- c('pk')

lista2 <- c()

lista_df_final <- list()

numero <- 1

if (existeAT == TRUE){
  tb_attest <- tb_attest %>% 
    rename(., Attest_Desc = diagnostico);
  lista1 <- append(lista1, 'Attest_Desc');
  lista_df_final[[1]] <- tb_attest;
  numero = numero + 1;
}

if (existeBL == TRUE){
 tb_block <- tb_block %>% 
   rename(., Block_Desc = diagnostico)
  lista1 <- append(lista1, 'Block_Desc');
  lista_df_final[[numero]] <- tb_block;
  numero = numero + 1;
}

if (existeEX == TRUE){
  tb_exam <- tb_exam %>% 
    rename(., Exam_Desc = descricao);
  lista1 <- append(lista1, 'Exam_Desc');
  lista_df_final[[numero]] <- tb_exam;
  numero = numero + 1;
  
}

if (existeP1 == TRUE){
  tb_presp1 <- tb_presp1 %>% 
  rename(., Presp1_Desc = receita);
  lista1 <- append(lista1, 'Presp1_Desc');
  lista_df_final[[numero]] <- tb_presp1;
  numero = numero + 1;
}

if (existeP3 == TRUE){
  tb_presp3 <- tb_presp3 %>% 
    rename(., Presp3_Desc = receita)
  lista1 <- append(lista1, 'Presp3_Desc')
  lista_df_final[[numero]] <- tb_presp3;
  numero = numero + 1;
}

if (existeIM == TRUE){
  tb_imagem <- tb_imagem %>% 
    rename(., Imagem = 'image')
  lista1 <- append(lista1, 'Imagem')
  lista_df_final[[numero]] <- tb_imagem;
}

# concatenar por id paciente

df_final <- lista_df_final %>% 
  reduce(full_join, by = "pk") %>% 
  select(lista1) %>% 
  mutate(Diagnostico = paste(Attest_Desc,
                             "\n",
                             Block_Desc,
                             "\n",
                             Presp1_Desc,
                             "\n",
                             Presp3_Desc,
                             "\n",
                             Exam_Desc,
                             "\n",
                             Imagem,
                             "\n",
                             sep="\n")) %>% 
  select(-Attest_Desc,-Block_Desc,-Presp1_Desc,-Presp3_Desc,-Exam_Desc,-Imagem) %>% 
  left_join(dadospacientes, by='pk')

###

df_final <- df_final %>% 
  mutate(Diagnostico = gsub("NA|(<p)(.*)(>)|<u>|<html>|<head>|</head>|<body>|</div>|</html>|</body>|</u>|<ul>|</ul>|<li>|<span>|</li>|</br>|<br>|<strong>|</strong>|</span>|<b>|</b>|\n\n\n\n\n\n\n\n\n|\n\n\n", " ", Diagnostico)) %>% 
  select(patient_id,patient_name,date,start_time,end_time,Diagnostico)

df_final <- df_final[!duplicated(df_final), ]

writexl::write_xlsx(df_final,'C:/Users/Doctoralia/Desktop/Doctoralia/Textos e scripts/Andrea 26-07-2022-event_record_TRATADO.xlsx')
