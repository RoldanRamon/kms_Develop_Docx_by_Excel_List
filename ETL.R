rm(list = ls())
library(dplyr)
library(readxl)
library(tidyr)
library(stringr)
library(purrr)
library(officer)

# Lê excel e separa por linhas cada modelo
base <- readxl::read_excel('./Arquivos/APP BOLETINS TECNICOS.xlsx',sheet = 'base') %>%
  janitor::clean_names() %>%
  tidyr::separate_rows(modelos, sep = ",\\s*") %>% 
  mutate(nome_do_arquivo = paste0(codigo,'_',modelos,'_') %>% stringr::str_replace_all(pattern = ' ',replacement = '_'))


# Função para criar um documento DOCX com uma linha como título e adicionar uma imagem
criar_docx <- function(row, codigo) {
  doc <- read_docx()
  doc <- doc %>%
    body_add_par(row, style = "heading 1") 
  
  caminho <- case_when(codigo == "fo" ~ "Imagens/Forwarder.jpeg",
                       codigo == "ha" ~ "Imagens/Harvester.JPG",
                       codigo == "pc" ~ "Imagens/PC_200.jpg",
                       codigo == "ca" ~ "Imagens/Cabecote.jpg",
                       codigo == "ll" ~ "Imagens/LogLoader.jpg",
                       codigo == "cs" ~ "Imagens/Guincho-Komatsu.jpg",
                       codigo == "tp" ~ "Imagens/TimberPro.png",
                       codigo == "fe" ~ "Imagens/TimberPro.png",
                       codigo == "si" ~ "Imagens/Silvicultura.png",
                       TRUE ~ ""
                       )
  
  if (caminho != "") {
    doc <- doc %>% 
      body_add_img(src = caminho, width = 8, height = 8)
  }
  
  
  return(doc)
}

# Mapear nomes de modelo diretamente
nomes_modelos <- base$nome_do_arquivo

# Iterar sobre cada linha e criar um documento DOCX para cada modelo
docs <- purrr::map2(nomes_modelos, base$codigo, criar_docx)

# Salvar cada documento em um arquivo separado com o nome do modelo
walk2(docs, nomes_modelos, ~print(.x, target = paste0("Exemplo de Boletins/", .y, "Exemplo_01.docx")))
walk2(docs, nomes_modelos, ~print(.x, target = paste0("Exemplo de Boletins/", .y, "Exemplo_02.docx")))
walk2(docs, nomes_modelos, ~print(.x, target = paste0("Exemplo de Boletins/", .y, "Exemplo_03.docx")))

