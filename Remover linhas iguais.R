rm(list = ls())

pacman::p_load('stringr', "tidyr", "dplyr", "readxl")

diretorio <- ""
df <- read_xlsx(diretorio)
df = data.frame(lapply(df, as.character), stringsAsFactors=FALSE)

df <- unique(df, fromLast = TRUE)


writexl::write_xlsx(df,diretorio)
