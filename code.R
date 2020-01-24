library(XML)
# library(xslt)
library(xlsx)
library(gWidgets)
# library(methods)
# library(officer)
# library(RDCOMClient)
library(zip)


setwd("U:\\Pastas pessoais\\Pedro\\Codigos e experimentos\\R\\Code_VBA_Chart")
# Arquivo vbs para criar charts
pathofvbscript <- "vba/VBA_Copy_Paste.vbs"

# extração

arquivo_original <- file.choose()

# Cria pastas
dir.create("docx_file",showWarnings = F)
dir.create("docx_extract",showWarnings = F)
dir.create("xlsx_file",showWarnings = F)
dir.create("xlsx_extract",showWarnings = F)

# Pasta de gráficos no Word
pasta_grafico_word <- "docx_extract/word/charts/"

# Copia o arquivo
file.copy(from = arquivo_original,to = "docx_file",overwrite = T)

# Extrai dados
nome_arquivo_docx <- list.files('docx_file/')
nome_arquivo_docx <- paste0("docx_file/",nome_arquivo_docx)

# Descompactar na pasta
unzip(zipfile = nome_arquivo_docx,exdir =  "docx_extract")

# Lista de arquivos
lista_arquivos <- list.files(path = "docx_extract/word/charts/")
lista_arquivos <- grep(pattern = "chart",x = lista_arquivos,value = T)

# Tira todas as tags referenciando a formulas
i = 1
for(charts in lista_arquivos){

  grafico <- xmlParseDoc(paste0(pasta_grafico_word,charts))
  
  numero_formulas <- length(grafico["//c:f"])
  # trycatch para ignorar os erros
  for(formula in 1:numero_formulas){
    tryCatch({
    xmlValue(grafico["//c:f"][[formula]])=""
    },error=function(e){})
  }
  
  #xmlValue(grafico["//c:f"][[2]]) <- ""
  
  #invisible(replaceNodes(grafico["//c:f"][[2]],newXMLTextNode(""))
  cat("\n",charts)
  saveXML(doc = grafico,file = paste0(pasta_grafico_word,charts))
  i = i+1
}

z7path = shQuote("U:/Pastas pessoais/Pedro/Codigos e experimentos/R/Code_VBA_Chart/7-Zip/7z.exe")
arquivo = 'docx_extract.docx'
setwd("docx_extract/")
lista_compactar <- list.files(path = ".",include.dirs = T,recursive = T,all.files = F)
# zipr(zipfile = "docx_extract.docx",
#     files = lista_compactar,
#     include_directories = F)
cmd = paste0(z7path," a -r ",arquivo)
shell(cmd)
setwd("..")

shell(shQuote(normalizePath(pathofvbscript)), "cscript", flag = "//nologo")


# salvar
# escolher_salvar <- file.choose(new = T)
# file.copy(from = "xlsx_file/xlsx_file_export.xlsx",to = escolher_salvar)

file_remove <- c(paste0("docx_extract/",list.files("docx_extract/",all.files = T,recursive = T)),
                 paste0("docx_file/",list.files("docx_file/",all.files = T,recursive = T)))

# file.copy(from = "xlsx_file/xlsx_file_export.xlsx",to = escolher_salvar)
file.remove(file_remove)

rm(list=ls())
