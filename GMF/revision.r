#Deben remplazar con dando ctrol + f 
#D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA

#por la dirección de las carpetas de su equipo. Recordar que es con /

#======================================== Parte 1 =================================
#-------------------------------- Instalación de paquetes 
#install.packages("googlesheets4")
#install.packages("tidyverse")
#install.packages("remotes")
#install.packages(c("FRACTION","dplyr","tidyverse","stringr","lubridate","tidyr","openxlsx","readxl","shiny","miniUI","timechange","taskscheduleR","openxlsx","writexl"))
#install.packages("stringi")  

#abrimos las librerias 
library(googledrive)
library(googlesheets4)
library(FRACTION)
library(dplyr)
library(stringr)
library(readxl)
library(shiny)
library(miniUI)
library(timechange)
library(lubridate)
library(tidyr)
library(openxlsx)
library(writexl)
library(readr)
library(stringi)

#Para que los datos no esten en anotación cientifica
options(scipen=999)


#------------------------------------------ Conceder permisos 
#autenticación pasar a google chrome para dar permisos al r 
gs4_auth()

#======================================== Parte 2 =====================
drive_auth()
#======================================== Parte 3 ==================================
#Importamos la estructura 
comprobante_1 = read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Estructura', col_names = TRUE, range = "A3:K144")
names(comprobante_1) = c("Consecutivo", "MASCARA", "MOTIVO","DESCRIPCION", "Numero_de_semana" ,"SEMANA","BASE","CALCULO_CONTR","CONTR_ASUMIDA","TOTAL_CONTRIBUCION_ME", "TOTAL_CONTRIBUCION_IMPUESTOS")

#Importamos la estructura de impuestos
comprobante_impuestos = read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Estructura_Impuestos', col_names = TRUE, range = "A1:BE117")
names(comprobante_impuestos) = c("CONSECUTIVO","LIBRO","ORIGEN","FECHA CONTABLE","DESCRIPCION","COMPANIA", "CUENTA CONTABLE","OFICINA","SUCURSAL","PROYECTO","SUBPROYECTO","TIPO COMPROBANTE","INTERCOMPANIA","VINCULADO","FUTURO1","FUTURO2","MONEDA", "DEBITO",                               
                                 "CREDITO","DESCRIPCION LINEA","REFERENCIA COMPLEMENTARIA 1","VALOR DE REFERENCIA COMPLEMENTARIA 1", "REFERENCIA COMPLEMENTARIA 2","VALOR DE REFERENCIA COMPLEMENTARIA 2", "REFERENCIA COMPLEMENTARIA 3","VALOR DE REFERENCIA COMPLEMENTARIA 3", "REFERENCIA COMPLEMENTARIA 4","VALOR DE REFERENCIA COMPLEMENTARIA 4", "REFERENCIA COMPLEMENTARIA 5","VALOR DE REFERENCIA COMPLEMENTARIA 5", "REFERENCIA COMPLEMENTARIA 6","VALOR DE REFERENCIA COMPLEMENTARIA 6", "REFERENCIA COMPLEMENTARIA 7","VALOR DE REFERENCIA COMPLEMENTARIA 7", 
                                 "REFERENCIA COMPLEMENTARIA 8","VALOR DE REFERENCIA COMPLEMENTARIA 8", "REFERENCIA COMPLEMENTARIA 9","VALOR DE REFERENCIA COMPLEMENTARIA 9", "REFERENCIA COMPLEMENTARIA 10","VALOR DE REFERENCIA COMPLEMENTARIA 10","REFERENCIA COMPLEMENTARIA 11","VALOR DE REFERENCIA COMPLEMENTARIA 11","REFERENCIA COMPLEMENTARIA 12","VALOR DE REFERENCIA COMPLEMENTARIA 12","REFERENCIA COMPLEMENTARIA 13","VALOR DE REFERENCIA COMPLEMENTARIA 13","REFERENCIA COMPLEMENTARIA 14","VALOR DE REFERENCIA COMPLEMENTARIA 14","REFERENCIA COMPLEMENTARIA 15","VALOR DE REFERENCIA COMPLEMENTARIA 15","REFERENCIA COMPLEMENTARIA 16","VALOR DE REFERENCIA COMPLEMENTARIA 16","REFERENCIA COMPLEMENTARIA 17","VALOR DE REFERENCIA COMPLEMENTARIA 17","REFERENCIA COMPLEMENTARIA 18", "VALOR DE REFERENCIA COMPLEMENTARIA 18","VALOR DE AUXILIAR")

#Importamos las fechas para hacer las semana 
mes = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B2"))
año = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D2"))

numero_semana_1 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "F6"))
numero_semana_2 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "F7"))
numero_semana_3 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "F8"))
numero_semana_4 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "F9"))
numero_semana_5 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "F10"))

#Semana 1 
De_1 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B6"))
De_Mes_1 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C6"))
Hasta_1 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D6"))
Hasta_Mes_1 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E6"))

#Semana 2
De_2 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B7"))
De_Mes_2 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C7"))
Hasta_2 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D7"))
Hasta_Mes_2 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E7"))

#Semana 3
De_3 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B8"))
De_Mes_3 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C8"))
Hasta_3 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D8"))
Hasta_Mes_3 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E8"))

#Semana 4
De_4 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B9"))
De_Mes_4 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C9"))
Hasta_4 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D9"))
Hasta_Mes_4 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E9"))

#Semana 5
De_5 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B10"))
De_Mes_5= as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C10"))
Hasta_5 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D10"))
Hasta_Mes_5 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E10"))

#Mes completo
De_6 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B11"))
De_Mes_6= as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "C11"))
Hasta_6 = as.numeric(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "D11"))
Hasta_Mes_6 = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "E11"))


#Parametros
parametros_cheques_girados = read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Parametro_Cheques_Girados', col_names = TRUE)
names(parametros_cheques_girados) = c("ID_COMPROBANTE", "ID_DE_ASIENTO", "Concepto")
parametros_cheques_girados = parametros_cheques_girados %>% mutate(ID_COMPROBANTE = as.character(ID_COMPROBANTE), 
                                                                   ID_DE_ASIENTO = as.character(ID_DE_ASIENTO),
                                                                   Concepto = as.character(Concepto))

#Colocamos las fechas de las semanas en el comprobante 
mes_letra  = case_when(mes == 1 ~ "Enero",mes == 2 ~ "Febrero", mes == 3 ~ "Marzo",mes == 4 ~ "Abril",mes == 5 ~ "Mayo",mes == 6 ~ "Junio",mes == 7 ~ "Julio",mes == 8 ~ "Agosto",mes == 9 ~ "Septiembre",mes == 10 ~ "Octubre",mes == 11 ~ "Noviembre",mes == 12 ~ "Diciembre",TRUE ~ "Error")
fecha_semana_1 = paste(De_1,"de",De_Mes_1,"al", Hasta_1, "de", Hasta_Mes_1 )
fecha_semana_2 = paste(De_2,"de",De_Mes_2,"al", Hasta_2, "de", Hasta_Mes_2)
fecha_semana_3 = paste(De_3,"de",De_Mes_3,"al", Hasta_3, "de", Hasta_Mes_3)
fecha_semana_4 = paste(De_4,"de",De_Mes_4,"al", Hasta_4, "de", Hasta_Mes_4)
fecha_semana_5 = paste(De_5,"de",De_Mes_5,"al", Hasta_5, "de", Hasta_Mes_5)
fecha_completa = paste(De_6,"de",De_Mes_6,"al", Hasta_6, "de", Hasta_Mes_6)

comprobante_1[c(1,9, 17, 25,33,41,49,57,65,73,81,89,97,105, 113, 121, 129),6] <- fecha_semana_1
comprobante_1[c(2,10, 18,26,34,42,50,58,66,74,82,90,98,106, 114, 122, 130),6] <- fecha_semana_2
comprobante_1[c(3,11, 19,27,35,43,51,59,67,75,83,91,99,107, 115, 123, 131),6] <- fecha_semana_3
comprobante_1[c(4,12, 20,28,36,44,52,60,68,76,84,92,100,108, 116, 124, 132),6] <- fecha_semana_4
comprobante_1[c(5,13, 21,29,37,45,53,61,69,77,85,93,101,109, 117, 125, 133),6] <- fecha_semana_5
comprobante_1[c(6,14, 22,30,38,46,54,62,70,78,86,94,102,110, 118, 126, 134),6] <- "Total"
comprobante_1[c(7,15, 23,31,39,47,55,63,71,79,87,95,103,111, 119, 127, 135),6] <- fecha_completa
comprobante_1[c(8,16, 24,32,40,48,56,64,72,80,88,96,104,112, 120, 128, 136),6] <- "Control"

#Limpiar el comprobante
comprobante_1 <- comprobante_1 %>% 
  mutate(across(c(7:11), ~ 0))


#=========================================== De drive al escritorio automaticamente =======================
# ID de la carpeta en Google Drive
folder_id = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B14"))

# Listar archivos en la carpeta
folder <- as_id(folder_id)
files <- drive_ls(path = folder)

# Nombres de archivos que quieres descargar
file_names <- c("CDAT_Semana_1.csv","CDAT_Semana_2.csv", "CDAT_Semana_3.csv","CDAT_Semana_4.csv","CDAT_Semana_5.csv","CDAT_Completo.csv",
                "CDT_Semana_1.csv","CDT_Semana_2.csv", "CDT_Semana_3.csv","CDT_Semana_4.csv","CDT_Semana_5.csv","CDT_Completo.csv",
                "Cheques_Oficina_Semana_1.xlsx","Cheques_Oficina_Semana_2.xlsx","Cheques_Oficina_Semana_3.xlsx","Cheques_Oficina_Semana_4.xlsx","Cheques_Oficina_Semana_5.xlsx","Cheques_Oficina_Completo.xlsx", 
                "Pago_Proveedores_Semana_1.xlsx","Pago_Proveedores_Semana_2.xlsx","Pago_Proveedores_Semana_3.xlsx","Pago_Proveedores_Semana_4.xlsx","Pago_Proveedores_Semana_5.xlsx","Pago_Proveedores_Completo.xlsx",
                "Cheques_Girados_Semana_1.xlsx", "Cheques_Girados_Semana_2.xlsx", "Cheques_Girados_Semana_3.xlsx", "Cheques_Girados_Semana_4.xlsx","Cheques_Girados_Semana_5.xlsx", "Cheques_Girados_Completo.xlsx",
                "Pago_Intereses_Semana_1.xlsx", "Pago_Intereses_Semana_2.xlsx","Pago_Intereses_Semana_3.xlsx","Pago_Intereses_Semana_4.xlsx","Pago_Intereses_Semana_5.xlsx","Pago_Intereses_Completo.xlsx", 
                "Timbre_Semana_1.xlsx", "Timbre_Semana_2.xlsx", "Timbre_Semana_3.xlsx", "Timbre_Semana_4.xlsx", "Timbre_Semana_5.xlsx", "Timbre_Completo.xlsx",
                "Redenciones_CDTS_Semana_1.DAT", "Redenciones_CDTS_Semana_2.DAT","Redenciones_CDTS_Semana_3.DAT","Redenciones_CDTS_Semana_4.DAT","Redenciones_CDTS_Semana_5.DAT","Redenciones_CDTS_Completo.DAT",
                "Redenciones_CDTS_Semana_1.TXT", "Redenciones_CDTS_Semana_2.TXT","Redenciones_CDTS_Semana_3.TXT","Redenciones_CDTS_Semana_4.TXT","Redenciones_CDTS_Semana_5.TXT","Redenciones_CDTS_Completo.TXT",
                "Fondo_Inversion_Semana_1.csv","Fondo_Inversion_Semana_2.csv","Fondo_Inversion_Semana_3.csv","Fondo_Inversion_Semana_4.csv","Fondo_Inversion_Semana_5.csv","Fondo_Inversion_Completo.csv")

# Directorio local donde deseas guardar los archivos descargados
local_directory <- "D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA"

#Eliminar los insumos de la corrida pasada
#unlink(local_directory, recursive = TRUE)

# Descargar los archivos por nombre
for (file_name in file_names) {
  file_to_download <- files[files$name == file_name,]
  if (nrow(file_to_download) > 0) {
    drive_download(file = file_to_download, path = file.path(local_directory, file_name), overwrite = TRUE)
  } else {
    cat("El archivo", file_name, "no se encontró en la carpeta de Google Drive.\n")
  }
}

#------------------------------------------ Importacion a R 
#definimos la parte fija --> cambiar en el escritorio de equipo CR -->  ojo con los / 
input = local_directory

#acordamos la dirección de entrada (input) de los archivos 
setwd(input)

#================================= GMF cheques y proveedores ===============================
Cheques_oficina = function(data,i) {
  tryCatch(
    {
      datos = read_excel(data, skip = 9) 
      colnames(datos) = c("Cuenta", "Descripcion_Cuenta" , "Sucursal" ,"Descripcion_Sucursal","Oficina","Descripcion_Oficina","Tipo_Comprobante", "Descripcion_Tipo_Comprobante", "Numero_de_Transaccion", "Descripcion_Transaccion","Fecha_Transaccion","Fecha_de_grabacion","Descripcion_Origen", "Numero_de_Identificacion","Digito_de_Verificacion","Razon_Social", "Documento", "Referencia_1", "Dependencia", "Subproyecto", "Auxiliar_de_Conciliacion", "Tipo_de_Evento", "Clase_Contable", "Numero_de_Linea", "Saldo_Dia_Anterior", "Debito", "Credito", "Asiento_Reversado") 
      datos = data.frame(datos) %>% 
        mutate(concepto = substr(Descripcion_Transaccion, start = 6, stop = 14),
               concepto3 = substr(concepto, start = 6, stop = 9),
               neto = Debito - Credito)
      
      datos = datos[,c("Fecha_Transaccion", "Descripcion_Transaccion","Debito","Credito", "Documento", "concepto", "neto", "Oficina", "concepto3", "Tipo_Comprobante")]
      
      #Llamamos para hacer un marge y que salga cheques con la fecha
      datos_cheques = read_excel(data, skip = 9)
      colnames(datos_cheques) = c("Cuenta", "Descripcion_Cuenta" , "Sucursal" ,"Descripcion_Sucursal","Oficina","Descripcion_Oficina","Tipo_Comprobante", "Descripcion_Tipo_Comprobante", "Numero_de_Transaccion", "Descripcion_Transaccion","Fecha_Transaccion","Fecha_de_grabacion","Descripcion_Origen", "Numero_de_Identificacion","Digito_de_Verificacion","Razon_Social", "Documento", "Referencia_1", "Dependencia", "Subproyecto", "Auxiliar_de_Conciliacion", "Tipo_de_Evento", "Clase_Contable", "Numero_de_Linea", "Saldo_Dia_Anterior", "Debito", "Credito", "Asiento_Reversado") 
      datos_cheques  = datos_cheques[,c("Oficina", "Fecha_Transaccion", "Descripcion_Transaccion", "Documento","Numero_de_Transaccion")]
      
      #-------------------------------  resultados 1  Cheques oficina 40-28-29 ------------------------ --------------
      #Filtramos por el concepto ='0098 0040'- '0099 0028'- '0099 0029'
      conceptos = filter(datos, datos$concepto == "0098 0040" | datos$concepto == "0099 0028" | datos$concepto == "0099 0029" )
      
      #Creamos la tabla con la sumatoria por concepto 
      las40 = filter(conceptos, conceptos$concepto == "0098 0040")
      las28 = filter(conceptos, conceptos$concepto == "0099 0028")
      las29 = filter(conceptos, conceptos$concepto == "0099 0029")
      
      las40_debito = sum(las40$Debito)
      las40_credito = sum(las40$Credito)
      
      las28_debito = sum(las28$Debito)
      las28_credito = sum(las28$Credito)
      
      las29_debito = sum(las29$Debito)
      las29_credito = sum(las29$Credito)
      Cuentas = c("Total 9710 Concepto 40", "Total 9610 Concepto 28","Total 9610 Concepto 29")
      debito = c(las40_debito, las28_debito, las29_debito)
      credito = c(las40_credito, las28_credito, las29_credito)
      
      resultado_cheques_oficina = data.frame(Cuentas, debito, credito)
      
      #Añadimos el total de todos los conceptos
      Cuentas = "Total concepto 40 + 28 +29"
      debito = sum(las40_debito, las28_debito, las29_debito)
      credito = sum(las40_credito, las28_credito, las29_credito)
      resultado_total = data.frame(Cuentas, debito, credito)
      
      resultado_cheques_oficina = rbind(resultado_cheques_oficina, resultado_total)
      resultado_cheques_oficina$neto = resultado_cheques_oficina$debito -  resultado_cheques_oficina$credito
      
      #añadimos el total de reposiciones
      Cuentas = "Total reposiciones 28 + 29"
      debito = sum(las28_debito, las29_debito)
      credito = sum(las28_credito, las29_credito)
      resultado_reposiciones = data.frame(Cuentas, debito, credito)
      
      resultado_reposiciones$neto = resultado_reposiciones$debito -  resultado_reposiciones$credito
      resultado_cheques_oficina = rbind(resultado_cheques_oficina, resultado_reposiciones)
      
      #---------------------------- Resultado 2 cheques de oficina Pagos ISA y Sobrantes FM  ------
      #NG - solo credito 
      #Pagos ISA - 0099 
      #Pagos sobrantes - 0053 
      
      #Filtamos por NG 
      isa_sobrantes = filter(datos, datos$Tipo_Comprobante == "NG")
      
      #Filtramos por 0099 en concepto 
      #Sum de credito 
      isa = filter(isa_sobrantes, isa_sobrantes$concepto3 == "0099")
      isa_credito = sum(isa$Credito)
      
      #Filtramos por 0053 en concepto 
      #Sum de credito 
      sobrantes = filter(isa_sobrantes, isa_sobrantes$concepto3 == "0053")
      sobrantes_credito = sum(sobrantes$Credito)
      
      #creamos la data 
      Concepto = c("Pagos ISA", "Sobrantes Fm")
      Credito = c(isa_credito, sobrantes_credito)
      
      resultado_sobrante_isa = data.frame(cbind(Concepto, Credito))
      
      #---------------------------- creamos el archivo cheques -------------------
      #extraemos los datos que se necesitan unicamente de la tabal "conceptos" que ya esta filtrada
      df = conceptos[,c("Fecha_Transaccion", "Descripcion_Transaccion","Documento", "neto", "Oficina")]
      
      prueba = data.frame(df %>%
                            mutate(s1 = sign(neto), absneto = abs(neto)) %>% 
                            arrange(Oficina, desc(absneto), s1 == -1) %>% 
                            group_by(Oficina,Fecha_Transaccion, Documento) %>% 
                            mutate(grp = cumsum(s1 > 0)) %>% 
                            group_by(grp, .add = TRUE) %>% 
                            filter(n_distinct(absneto) > 1|n_distinct(s1) == 1) %>% 
                            group_by(s1, .add = TRUE) %>% 
                            summarise(neto = sum(neto), .groups = "drop") %>%   
                            select(-grp, -s1))
      
      cheques = data.frame (prueba %>%
                              mutate(abs_neto=abs(neto)) %>%
                              group_by(Oficina, Fecha_Transaccion, abs_neto) %>%
                              mutate(one_plus_one_minus= any(neto>=0) & any(neto<=0))%>%
                              filter(n()==1 | !one_plus_one_minus) %>%
                              ungroup %>%
                              select(- abs_neto, - one_plus_one_minus))
      
      
      #merge con el df para traer los otros datos 
      cheques = merge(cheques, df, by = c("Oficina", "Fecha_Transaccion", "Documento" ,"neto"))
      cheques = cheques[,c("Oficina" , "Fecha_Transaccion" , "Descripcion_Transaccion", "Documento", "neto")]
      
      cheques = left_join(cheques, datos_cheques, by = c("Oficina", "Fecha_Transaccion", "Descripcion_Transaccion", "Documento"))
      cheques = cheques[,c("Numero_de_Transaccion", "Oficina","Fecha_Transaccion","Descripcion_Transaccion", "Documento","neto")]
      
      #Aqui agregamos lo que faltaba en calculos del GMF para no tocar el codigo más arriba de lo que ya se tenia
      resultado_cheques_oficina = resultado_cheques_oficina %>% mutate(
        neto =abs(neto),
        GMF = neto * 0.004
      )
      resultado_cheques_oficina[5,4]  = resultado_cheques_oficina[5,4] * -1
      resultado_cheques_oficina[5,5]  = resultado_cheques_oficina[5,5] * -1
      
      resultado_sobrante_isa = resultado_sobrante_isa %>% mutate(
        Credito =as.numeric(Credito),
        Credito =abs(Credito),
        GMF = Credito * 0.004
      )
      
      #Calculamos los cheques mayores a COP 15.000.000 
      cheques_mayor_a = filter(cheques, abs(cheques$neto) >= 15000000)
      
      #Exportamos al ambiente globlal 
      #Creamos indices
      index <- i
      var_name_1 <- paste("R_cheques_oficina_", index, sep = "")
      var_name_2 <- paste("R_ISA_sobrantes_", index, sep = "")
      var_name_3 <- paste("Cheques_", index, sep = "")
      var_name_4 <- paste("Nombre_R_cheques_", index, sep = "")
      var_name_5 <- paste("Nombre_R_ISA_", index, sep = "")
      var_name_6 <- paste("Cheques_Oficina_", index, sep = "")

      assign(var_name_1, resultado_cheques_oficina  , envir = .GlobalEnv)
      assign(var_name_2, resultado_sobrante_isa, envir = .GlobalEnv)
      assign(var_name_3, cheques, envir = .GlobalEnv)
      assign(var_name_4, var_name_1, envir = .GlobalEnv)
      assign(var_name_5, var_name_2, envir = .GlobalEnv)
      assign(var_name_6, cheques_mayor_a, envir = .GlobalEnv)

      
    },
    error = function(e){ 
      mensaje = paste('El archivo', data, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}
Cheques_Oficina_Semana_1 = "Cheques_Oficina_Semana_1.xlsx"
Cheques_Oficina_Semana_2 = "Cheques_Oficina_Semana_2.xlsx"
Cheques_Oficina_Semana_3 = "Cheques_Oficina_Semana_3.xlsx"
Cheques_Oficina_Semana_4 = "Cheques_Oficina_Semana_4.xlsx"
Cheques_Oficina_Semana_5 = "Cheques_Oficina_Semana_5.xlsx"
Cheques_Oficina_Completo = "Cheques_Oficina_Completo.xlsx"

Cheques_oficina (Cheques_Oficina_Semana_1, 1)
Cheques_oficina (Cheques_Oficina_Semana_2, 2)
Cheques_oficina (Cheques_Oficina_Semana_3, 3)
Cheques_oficina (Cheques_Oficina_Semana_4, 4)
Cheques_oficina (Cheques_Oficina_Semana_5, 5)
Cheques_oficina (Cheques_Oficina_Completo, 6)


#Funciones pago provedores
Pago_proveedores = function(datos, i) {
  tryCatch(
    {
      datos = read_excel(datos, skip = 9) 
      colnames(datos) = c("Cuenta", "Descripcion_Cuenta" , "Sucursal" , "Descripcion_Sucursal","Oficina","Descripcion_Oficina", "Tipo_Comprobante", "Descripcion_Tipo_Comprobante", "Numero_de_Transaccion","Descripcion_Transaccion","Fecha_Transaccion","Fecha_de_grabacion","Descripcion_Origen", "Numero_de_Identificacion","Digito_de_Verificacion","Razon_Social", "Documento", "Referencia_1","Dependencia", "Subproyecto", "Auxiliar_de_Conciliacion", "Tipo_de_Evento", "Clase_Contable", "Numero_de_Linea", "Saldo_Dia_Anterior", "Debito", "Credito", "Asiento_Reversado")  
      datos = data.frame(datos) 

      AH = filter(datos, datos$Tipo_Comprobante == "AH")
      FD = filter(datos, datos$Tipo_Comprobante == "FD")
      MB = filter(datos, datos$Tipo_Comprobante == "MB")
      
      AH_debito = sum(AH$Debito)
      AH_credito = sum(AH$Credito)
      
      FD_debito = sum(FD$Debito)
      FD_credito = sum(FD$Credito)
      
      MB_debito = sum(MB$Debito)
      MB_credito = sum(MB$Credito)
      
      
      Cuentas = c("Pago proveedores Damas","Pago proveedores Fijo Diario","Pago proveedores Cuenta Corriente")
      Descripcion = c("AH", "FD", "MB")
      Debito = c(AH_debito, FD_debito, MB_debito)
      Credito = c(AH_credito, FD_credito, MB_credito)
      
      resultado_pago_proveedores = data.frame(Cuentas, Descripcion, Debito, Credito)
      
      #Aqui agregamos lo que faltaba en calculos del GMF para no tocar el codigo más arriba de lo que ya se tenia
      resultado_pago_proveedores = resultado_pago_proveedores %>% mutate(
        neto = abs (Debito + Credito),
        GMF = neto * 0.004
      )
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("R_pago_proveedores_", index, sep = "")
      var_name_2 <- paste("Nombre_R_Provedores_", index, sep = "")
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, resultado_pago_proveedores, envir = .GlobalEnv)
      assign(var_name_2, var_name_1, envir = .GlobalEnv)
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}

Pago_Proveedores_Semana_1 = "Pago_Proveedores_Semana_1.xlsx"
Pago_Proveedores_Semana_2 = "Pago_Proveedores_Semana_2.xlsx"
Pago_Proveedores_Semana_3 = "Pago_Proveedores_Semana_3.xlsx"
Pago_Proveedores_Semana_4 = "Pago_Proveedores_Semana_4.xlsx"
Pago_Proveedores_Semana_5 = "Pago_Proveedores_Semana_5.xlsx"
Pago_Proveedores_Completo = "Pago_Proveedores_Completo.xlsx"

Pago_proveedores(Pago_Proveedores_Semana_1, 1) 
Pago_proveedores(Pago_Proveedores_Semana_2, 2) 
Pago_proveedores(Pago_Proveedores_Semana_3, 3) 
Pago_proveedores(Pago_Proveedores_Semana_4, 4) 
Pago_proveedores(Pago_Proveedores_Semana_5, 5) 
Pago_proveedores(Pago_Proveedores_Completo, 6)
#================================= GMF - CDT - CDAT ===============================

#Función CDAT 
CDAT = function(datos, i) {
  tryCatch(
    {
      
      datos = read_csv (datos, col_types = cols_only(op_num_banco3 = col_character(), rpt_tran = col_character(), rpt_valor = col_character(), rpt_nombre_cliente = col_character(), rpt_concepto = col_character(), rpt_num_id = col_character(), en_nomlar2 = col_character(), rpt_tip_id = col_number()), skip = 2,locale = readr::locale(encoding = "latin1"))
      
      datos$rpt_valor <- parse_number(gsub("[^0-9.-]", "", datos$rpt_valor)) * 
        ifelse(grepl("\\(", datos$rpt_valor), -1, 1)
      
      datos = data.frame(datos) 
      
      #Creamos los grupos 
      G1_CDAT = filter(datos, (datos$rpt_tran == 14943 & datos$rpt_concepto == 'PAGO DE INTERESES Y OTROS'))
      
      #Aqui agregamos lo que faltaba en calculos del GMF para no tocar el codigo más arriba de lo que ya se tenia
      resumen_G1_CDAT = data.frame(Total = sum(G1_CDAT$rpt_valor)*-1,
                                   GMF = sum((G1_CDAT$rpt_valor) * 0.004)*-1)
      
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("resumen_CDAT_", index, sep = "")
      var_name_2 <- paste("Nombre_R_CDAT_", index, sep = "")
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, resumen_G1_CDAT, envir = .GlobalEnv)
      assign(var_name_2, var_name_1, envir = .GlobalEnv)
      
      
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}

CDAT_Semana_1 = "CDAT_Semana_1.csv"
CDAT_Semana_2 = "CDAT_Semana_2.csv"
CDAT_Semana_3 = "CDAT_Semana_3.csv"
CDAT_Semana_4 = "CDAT_Semana_4.csv"
CDAT_Semana_5 = "CDAT_Semana_5.csv"
CDAT_Completo = "CDAT_Completo.csv"

CDAT(CDAT_Semana_1, 1) 
CDAT(CDAT_Semana_2, 2) 
CDAT(CDAT_Semana_3, 3) 
CDAT(CDAT_Semana_4, 4) 
CDAT(CDAT_Semana_5, 5) 
CDAT(CDAT_Completo, 6) 

#Función CDT 
CDT = function(datos, i,fondo_de_inversion, redenciones_individuales) {
  tryCatch(
    {
      
      datos <- read_csv(datos, col_types = cols(pd_fecha_proceso = col_date(format = "%d/%m/%Y"), 
                                                rpt_fecha_aplicacion = col_date(format = "%d/%m/%Y"), 
                                                rpt_num_id = col_character(),
                                                op_num_banco3 = col_character(),
                                                rpt_tran = col_character(), 
                                                rpt_valor = col_character(),
                                                rpt_nombre_cliente = col_character(), 
                                                rpt_concepto = col_character(),
                                                rpt_num_id = col_character(), 
                                                en_nomlar2 = col_character(),
                                                rpt_tip_id = col_number()), skip = 2,locale = readr::locale(encoding = "latin1"))
      
      
      #Pasamos rpt_valor a dato numerico
      datos$rpt_valor <- parse_number(gsub("[^0-9.-]", "", datos$rpt_valor)) * 
        ifelse(grepl("\\(", datos$rpt_valor), -1, 1)
      
      datos = data.frame(datos) 
      
      # ======================================== Cruzamos con redenciones individuales =============
      #Llamamos el archivo de redenciones
      tryCatch( {
        
        
        datos_redenciones <- read_delim(redenciones_individuales, delim = ";", escape_double = FALSE, col_types = cols(NRO. ID COBIS. = col_character(), 
                                                                                                                       NRO CDT/CDTAS. = col_character(), 
                                                                                                                       VAL REDENCION CAPITAL = col_character(), 
                                                                                                                       VALOR GMF = col_character(), 
                                                                                                                       FEC PAGO = col_date(format = "%Y%m%d"),
                                                                                                                       NRO. ID CTA  = col_character(),
                                                                                                                       RECHAZOS Y OBSERVACIONES = col_character()),  trim_ws = TRUE, skip = 1)
        names(datos_redenciones) = c("TI_1", "NRO_ID_COBIS","NOMBRE_O_RAZON_SOCIAL_COBIS", "TIPO_MANEJO_COBIS","NRO_CDT_CDTAS","VAL_REDENCION_CAPITAL","NRO_CTA_ABONO","VALOR_GMF","NRO_CTA_HOMOLOGA","TIPO_DE_MANEJO_CTA","TI_11", "NRO_ID_CTA","NOMBRE_O_RAZON_SOCIAL_CTA","VALOR_ABONO_A_CUENTA","FEC_PAGO","RECHAZOS_Y_OBSERVACIONES") 
        datos_redenciones = datos_redenciones %>% mutate(VALOR_GMF = str_replace_all(VALOR_GMF,"[$]", ""),
                                                         VALOR_GMF = str_replace_all(VALOR_GMF,"[.]", ""),
                                                         VALOR_GMF = str_replace_all(VALOR_GMF,"[,]", "."),
                                                         VALOR_GMF = as.numeric(VALOR_GMF), 
                                                         
                                                         VAL_REDENCION_CAPITAL = str_replace_all(VAL_REDENCION_CAPITAL,"[$]", ""),
                                                         VAL_REDENCION_CAPITAL = str_replace_all(VAL_REDENCION_CAPITAL,"[.]", ""),
                                                         VAL_REDENCION_CAPITAL = str_replace_all(VAL_REDENCION_CAPITAL,"[,]", "."),
                                                         VAL_REDENCION_CAPITAL = as.numeric(VAL_REDENCION_CAPITAL)
        )
        
        datos_AF_AE = filter(datos_redenciones, substr(datos_redenciones$NRO_CDT_CDTAS, 5,6) == "AF" | substr(datos_redenciones$NRO_CDT_CDTAS, 5,6) == "AE" )
        datos_redenciones = anti_join(datos_redenciones, datos_AF_AE)
        datos_redenciones = data.frame(datos_redenciones)
        datos_total = filter(datos_redenciones, datos_redenciones$NRO_CDT_CDTAS == "TOTALES -->") 
        datos_redenciones = anti_join(datos_redenciones, datos_total)
        
        #Convertimos la columna NRO_ID_CTA en número para eliminar los ceros y redondeamos el valor
        datos_redenciones = datos_redenciones %>% mutate(
          NRO_ID_CTA = as.numeric(NRO_ID_CTA),
          NRO_ID_CTA = as.character(NRO_ID_CTA),
          VAL_REDENCION_CAPITAL = round(VAL_REDENCION_CAPITAL, 0)
        )
        
        #Agregarle una columna como llave - concatenamo: NRO_ID_CTA + NRO_CDT_CDTAS + VAL_REDENCION_CAPITAL
        datos_redenciones$llave = paste0(datos_redenciones$NRO_ID_CTA, datos_redenciones$NRO_CDT_CDTAS, datos_redenciones$VAL_REDENCION_CAPITAL)
        
        
        #-------------Pendiente de saber si primero quitamos deceval e intereses
        
        #Como los datos de redenciones vienen en valor absoluto, se pasa una columna de datos en absoluto
        datos$rpt_valor_absoluto = round(abs(datos$rpt_valor), 0)
        
        #Si esta el archivo vamos a cruzar con  los datos de rtram - monetaria
        datos$llave = paste0(datos$rpt_num_id, datos$op_num_banco3, datos$rpt_valor_absoluto)
        
        
        #Merge por izquierda para cruzar los datos que estan en las dos bases
        cruce_general = full_join(datos, datos_redenciones, by = "llave")
        
        #Quitamos todo lo que cruza para el control
        cruza_rtran_redenciones <- cruce_general %>%
          filter(!is.na(NRO_CDT_CDTAS))
        
        no_cruzan_rtran_redenciones <- cruce_general %>%
          filter(is.na(NRO_CDT_CDTAS))
        
        #Aquí sacamos la base que cruzo para el control semanal del CDT
        cruza_rtran_redenciones = data.frame(cruza_rtran_redenciones)
        
        cruza_rtran_redenciones <- cruza_rtran_redenciones %>%
          mutate(across(where(is.character), ~ stri_enc_toutf8(., validate = TRUE)))
        
        index <-i
        var_name_redencion <- paste("Cruza_rtran_redenciones_", index, sep = "")
        assign(var_name_redencion, cruza_rtran_redenciones, envir = .GlobalEnv)
        
        #Quitamos la ultima columna de llave que agregamos al rtran monetaria
        cruza_rtran_redenciones = cruza_rtran_redenciones %>% select(-c("rpt_valor_absoluto","llave","TI_1","NRO_ID_COBIS","NOMBRE_O_RAZON_SOCIAL_COBIS","TIPO_MANEJO_COBIS","NRO_CDT_CDTAS",              
                                                                        "VAL_REDENCION_CAPITAL","NRO_CTA_ABONO","VALOR_GMF","NRO_CTA_HOMOLOGA","TIPO_DE_MANEJO_CTA","TI_11",                      
                                                                        "NRO_ID_CTA", "NOMBRE_O_RAZON_SOCIAL_CTA","VALOR_ABONO_A_CUENTA","FEC_PAGO","RECHAZOS_Y_OBSERVACIONES" 
        ))
        
        #Quitamos llave y valor absolito de datos 
        datos = datos %>% select(-c("rpt_valor_absoluto","llave"))
        
        #Hacemos un antijoin para dejar los datos que no cruzaron para continuar con el resto de cruces
        datos = anti_join(datos, cruza_rtran_redenciones)
        
      }, 
      error = function(e){ 
        mensaje = paste('El archivo', redenciones_individuales, 'no esta disponible')
        cat(mensaje, "\n")
      }
      )
#============================================ CDT =========================================
      #Filtramos quitando Depositos y banco davivienda
      DECEVAL_DAVIVIENDA = filter (datos, datos$rpt_num_id == "8001820912" | datos$rpt_num_id == "8600343137")
      datos_sin_deceval_davivienda = filter (datos, datos$rpt_num_id != "8001820912" & datos$rpt_num_id != "8600343137")
      
      #Totalizamos Deceval
      DECEVAL = filter (DECEVAL_DAVIVIENDA, DECEVAL_DAVIVIENDA$en_nomlar2 != 'SEB' & DECEVAL_DAVIVIENDA$rpt_num_id != "8600343137")
      DECEVAL = filter (DECEVAL, DECEVAL$rpt_concepto == "PAGO DE INTERESES Y OTROS")
      total_cdt_deceval = sum(DECEVAL$rpt_valor)  #Enviar resumen
      
      #Creamos los usb grupos por cuenta y concepto 
      G1_CDT = filter(datos_sin_deceval_davivienda, (datos_sin_deceval_davivienda$rpt_tran == 14543 & datos_sin_deceval_davivienda$rpt_concepto == 'ANULACION POR CHEQUE DEVUELTO') |
                        (datos_sin_deceval_davivienda$rpt_tran == 14875 & datos_sin_deceval_davivienda$rpt_concepto == 'ANULACION DE PLAZO FIJO')|
                        (datos_sin_deceval_davivienda$rpt_tran == 14943 & datos_sin_deceval_davivienda$rpt_concepto == 'DEVOLUCION REMANENTE POR CHEQUE DEVUELTO'))
      
      G2_CDT = filter(datos_sin_deceval_davivienda, datos_sin_deceval_davivienda$rpt_tran == 14901 & datos_sin_deceval_davivienda$rpt_concepto == 'APERTURA DPF')
      
      G3_CDT = filter(datos_sin_deceval_davivienda, (datos_sin_deceval_davivienda$rpt_tran == 14903 & datos_sin_deceval_davivienda$rpt_concepto == 'CANCELACION DE PLAZO FIJO'))
      G4_CDT = filter(datos_sin_deceval_davivienda, (datos_sin_deceval_davivienda$rpt_tran == 14919 & datos_sin_deceval_davivienda$rpt_concepto == 'APERTURA POR REINVERSION') | (datos_sin_deceval_davivienda$rpt_tran == 14919 & datos_sin_deceval_davivienda$rpt_concepto == 'CANCELACION POR REINVERSION'))
      
      Intereses = filter(datos_sin_deceval_davivienda, datos_sin_deceval_davivienda$rpt_tran == 14943 & datos_sin_deceval_davivienda$rpt_concepto == 'PAGO DE INTERESES Y OTROS') 
      total_14943_Pago_Intereses_otros = sum(Intereses$rpt_valor) #Enviar resumen
      
      #============================= Primera parte: G1 vs G2 ===================================
      #Creamos una data organizando los grupo 1vs2
      CRUCE_1_CDT = rbind(G1_CDT, G2_CDT)
      
      #Haremos un group_by para totalizar todo lo que tenga coincidente en: op_num_banco3, rpt_num_id, rpt_tran,rpt_nombre_cliente,rpt_concepto
      Primera_agrupacion_cruce_1 = CRUCE_1_CDT %>% 
        group_by(op_num_banco3, rpt_num_id, rpt_tran,rpt_nombre_cliente,rpt_concepto, rpt_tip_id) %>% summarise(rpt_valor = sum(rpt_valor))
      
      #Añadimos la columna 'codigo' que pondra todo lo diferente a las 14901 con codigo -14901, para que me permita filtrar.
      Primera_agrupacion_cruce_1$codigo = case_when(Primera_agrupacion_cruce_1$rpt_tran == 14543 | Primera_agrupacion_cruce_1$rpt_tran == 14875 | Primera_agrupacion_cruce_1$rpt_tran == 14943 ~ -14901, TRUE ~ 14901)
      
      #======================== Primera limpieza: Revisar que se puede cancelar teniendo en cuenta por 'op_num_banco3' 
      muertos1 <- Primera_agrupacion_cruce_1 %>%
        semi_join(Primera_agrupacion_cruce_1 %>%
                    mutate(rpt_valor = -rpt_valor, codigo = -codigo), 
                  by = c("rpt_nombre_cliente", "rpt_num_id", "op_num_banco3","rpt_valor", "codigo"))
      
      muertos1_completo = merge(muertos1, datos_sin_deceval_davivienda, by = c("op_num_banco3","rpt_num_id","rpt_tran","rpt_nombre_cliente","rpt_concepto", "rpt_tip_id"))
      muertos1_completo =  muertos1_completo %>% mutate(rpt_valor = rpt_valor.x) %>% select(-c(codigo, rpt_valor.x, rpt_valor.y))
      orden = names(datos_sin_deceval_davivienda)
      muertos1_completo = muertos1_completo[,orden]
      
      # Obtener la tabla "vivos1" con las filas que no tienen contrapartes
      vivos1 <- Primera_agrupacion_cruce_1 %>%
        anti_join(muertos1, by = c("op_num_banco3", "rpt_tran", "rpt_valor", "rpt_nombre_cliente", "rpt_concepto", "rpt_num_id", "codigo"))
      
      
      #======================== Segunda limpieza: Revisar que se puede cancelar sin tener en cuenta 'op_num_banco3'
      #Sobre vivos1 vamos a hacer una segunda limpieza  sin tener en cuenta op_num_banco3 para ver que se puede cancelar o tiene contraparte
      muertos2 <- vivos1 %>%
        semi_join(vivos1 %>%
                    mutate(rpt_valor = -rpt_valor, codigo = -codigo), 
                  by = c("rpt_nombre_cliente", "rpt_num_id", "rpt_valor", "codigo"))
      
      muertos2_completo = merge(muertos2, datos_sin_deceval_davivienda, by = c("rpt_num_id","rpt_tran","rpt_nombre_cliente","rpt_concepto", "rpt_tip_id"))
      muertos2_completo = muertos2_completo %>% mutate(op_num_banco3 = op_num_banco3.x,
                                                       rpt_valor = rpt_valor.x) %>% select(-c(codigo, op_num_banco3.x,  op_num_banco3.y, rpt_valor.x, rpt_valor.y))
      
      muertos2_completo = muertos2_completo[,orden]
      
      # Obtener la tabla "vivos1" con las filas que no tienen contrapartes
      vivos2 <- vivos1 %>%
        anti_join(muertos2, by = c("op_num_banco3", "rpt_tran", "rpt_valor", "rpt_nombre_cliente", "rpt_concepto", "rpt_num_id", "codigo"))
      
      #========================= Tercera limpieza: Desagrupar para ver que podemos cancelar nuevamente  
      #Hacemos un merge por "op_num_banco3","rpt_num_id",  "rpt_tran", "rpt_nombre_cliente", "rpt_concepto" -- así en 'rpt_valor.y' tendremos los datos originales sin agrupar para ver que podemos cancelar sin tener en cuenta 'op_num_banco3'
      desagrupado_vivo <- merge(vivos2, CRUCE_1_CDT, by = c("op_num_banco3","rpt_num_id",  "rpt_tran", "rpt_nombre_cliente", "rpt_concepto"))
      
      #Volvimos a limpiar muertos3 y vivos3 sin tener en cuenta 'op_num_banco3' 
      muertos3 <- desagrupado_vivo %>%
        semi_join(desagrupado_vivo %>%
                    mutate(rpt_valor.y = -rpt_valor.y, codigo = -codigo), 
                  by = c("rpt_nombre_cliente", "rpt_num_id", "rpt_valor.y", "codigo"))
      
      #Finalmente es la tabla que queda ya con las operaciones vivas
      vivos3 <- desagrupado_vivo %>%
        anti_join(muertos3, by = c("op_num_banco3", "rpt_tran", "rpt_valor.y", "rpt_nombre_cliente", "rpt_concepto", "rpt_num_id", "codigo"))
      
      
      #Unimos los 3 que quedan muertos, quitando las columnas que no usamos para exportar un 'muertos_final_cruce_1_CDT' 
      muertos1 = muertos1 %>% select(-c('codigo'))
      muertos2 = muertos2 %>% select(-c('codigo'))
      muertos3 = muertos3 %>% mutate(rpt_valor = rpt_valor.y,
                                     rpt_tip_id = rpt_tip_id.x) %>% select(-c('codigo', 'rpt_valor.x', 'rpt_valor.y', "rpt_tip_id.x","rpt_tip_id.y"))
      
      muertos_final_cruce_1_CDT = rbind(muertos1, muertos2, muertos3) #Este se exporta 
      muertos_3_completo = muertos3[,orden]
      muertos_final_completo_CRUCE_1 = rbind(muertos1_completo, muertos2_completo, muertos_3_completo)
      
      
      #Ahora limpiamos vivos3 y separamos por codigos 
      vivos3 = vivos3 %>% mutate(rpt_valor = rpt_valor.y,
                                 rpt_tip_id= rpt_tip_id.x) %>% select(-c('codigo', 'rpt_valor.x', 'rpt_valor.y',"rpt_tip_id.x","rpt_tip_id.y"))
      
      #Hacemos vivos_G1_CDT y vivos_G2_CDT
      vivos_G1_CDT = filter(vivos3, vivos3$rpt_tran == 14543 | vivos3$rpt_tran == 14875 | vivos3$rpt_tran == 14943)
      vivos_G2_CDT = filter(vivos3, vivos3$rpt_tran == 14901)
      
      
      #============================= Segunda parte: G3 vs G4 ===================================
      Tabla_dinamica_G3  = data.frame(G3_CDT %>% group_by(rpt_num_id) %>% summarise(valor = sum(rpt_valor)))
      Tabla_dinamica_G4  = data.frame(G4_CDT %>% group_by(rpt_num_id) %>% summarise(valor = sum(rpt_valor)))
#unimos las dos 
      Tabla_dinamica_full = rbind(Tabla_dinamica_G3,Tabla_dinamica_G4)
      Tabla_dinamica_full  = data.frame(Tabla_dinamica_full %>% group_by(rpt_num_id) %>% summarise(valor = sum(valor)))
      
      Vivos_G3_vs_G4 = filter(Tabla_dinamica_full, Tabla_dinamica_full$valor != 0)
      muertos_G3_VS_G4 = filter(Tabla_dinamica_full, Tabla_dinamica_full$valor == 0)
      Vivos_G3_vs_G4 = anti_join(Vivos_G3_vs_G4, muertos_G3_VS_G4)
      
      #Desagrupamos los que se murieron 
      muertos_G3 = muertos_G3_VS_G4[,c(1:2)]
      muertos_G3_completo = merge(muertos_G3, G3_CDT, by = "rpt_num_id")
      muertos_G3_completo = muertos_G3_completo %>% select(-valor)
      muertos_G3_completo = muertos_G3_completo[,orden]
      
      muertos_G4 = muertos_G3_VS_G4[,c(1:2)]
      muertos_G4_completo = merge(muertos_G4, G4_CDT, by = "rpt_num_id")
      muertos_G4_completo = muertos_G4_completo %>% select(-valor)
      muertos_G4_completo = muertos_G4_completo[,orden]
      
      Muertos_completo_G3_VS_G4 = rbind(muertos_G3_completo,muertos_G4_completo)
      
      #Ahora sacamos los vivmos del grupo 3 y grupo 4 
      VIVOS_G3_CDT = anti_join(G3_CDT, muertos_G3_completo)
      VIVOS_G4_CDT = anti_join(G4_CDT, muertos_G4_completo)
      
      
      #============================= Tercera parte parte: vivos_G2_CDT_2 vs vivos_G4_CDT ===================================
      Tabla_dinamica_G2_4  = data.frame(vivos_G2_CDT %>% group_by(rpt_num_id) %>% summarise(valor = sum(rpt_valor)))
      Tabla_dinamica_G4_4  = data.frame(VIVOS_G4_CDT %>% group_by(rpt_num_id) %>% summarise(valor = sum(rpt_valor)))
      
      #unimos las dos 
      Tabla_dinamica_full_tercera_parte = rbind(Tabla_dinamica_G2_4,Tabla_dinamica_G4_4)
      Tabla_dinamica_full_tercera_parte  = data.frame(Tabla_dinamica_full_tercera_parte %>% group_by(rpt_num_id) %>% summarise(valor = sum(valor)))
      
      Vivos_G2_vs_G4 = filter(Tabla_dinamica_full_tercera_parte, Tabla_dinamica_full_tercera_parte$valor != 0)
      muertos_G2_VS_G4 = filter(Tabla_dinamica_full_tercera_parte, Tabla_dinamica_full_tercera_parte$valor == 0)
      Vivos_G2_vs_G4 = anti_join(Vivos_G2_vs_G4, muertos_G2_VS_G4)
      
      #Desagrupamos los que se murieron 
      muertos_G2_4 = muertos_G2_VS_G4[,c(1:2)]
      muertos_G2_completo_4 = merge(muertos_G2_4, vivos_G2_CDT, by = "rpt_num_id")
      muertos_G2_completo_4 = muertos_G2_completo_4 %>% select(-valor)
      muertos_G2_completo_4 = muertos_G2_completo_4[,orden]
      
      muertos_G4_4 = muertos_G2_VS_G4[,c(1:2)]
      muertos_G4_completo_4 = merge(muertos_G4_4, VIVOS_G4_CDT, by = "rpt_num_id")
      muertos_G4_completo_4 = muertos_G4_completo_4 %>% select(-valor)
      muertos_G4_completo_4 = muertos_G4_completo_4[,orden]
      
      Muertos_completo_G2_VS_G4_4 = rbind(muertos_G2_completo_4,muertos_G4_completo_4)
      
      #Ahora sacamos los vivmos del grupo 3 y grupo 4 
      VIVOS_G2_CDT_4 = anti_join(vivos_G2_CDT, muertos_G2_completo_4)
      VIVOS_G4_CDT_4 = anti_join(VIVOS_G4_CDT, muertos_G4_completo_4)
      
      
      #============================= Cruce fondo de inversión VS 14903 ===========
      tryCatch( {
        # Llamamos el archivo de fondo de inversion y lo limpiamos
        Fondo_Inversion <- read_csv(fondo_de_inversion, 
                                    col_types = cols(COD_TIPO_IDENTIFICACION = col_character(), #1 
                                                     NRO_IDENTIFICACION = col_character(),  #2
                                                     NRO_CUENTA = col_character(), #3
                                                     TIPO_FONDO = col_character(), #4
                                                     COD_FONDO = col_character(),  #5
                                                     TIPO_FLUJO = col_character(), #6
                                                     TIPO_TRANSACCION = col_character(), #7 
                                                     ESTADO_TRANSACCION = col_character(),#8
                                                     MONTO_NETO = col_double(), #9
                                                     MONTO_BRUTO = col_double(), #10
                                                     NOMBRE_COMPANIA = col_character(), #11
                                                     PERIODO = col_date(format = "%Y-%m-%d"))) #12
        
        # Limpieza de fondo de inversión
        #Fondo_Inversion$llave = paste0(Fondo_Inversion$NRO_IDENTIFICACION,Fondo_Inversion$TIPO_TRANSACCION)
        Fondo_Inversion = Fondo_Inversion[,c(2,4,7,9,12)]
        names(Fondo_Inversion)= c("rpt_num_id","tipo_fondo","tipo_transaccion","rpt_valor_absoluto","rpt_fecha_aplicacion")
        
        VIVOS_G3_CDT_fondo_inversion = VIVOS_G3_CDT
        
        VIVOS_G3_CDT_fondo_inversion <- VIVOS_G3_CDT_fondo_inversion %>%
          mutate(
            rpt_fecha_aplicacion = as.Date(rpt_fecha_aplicacion, format = "%Y-%m-%d"),
            rpt_valor_absoluto = round(abs(rpt_valor), 0),
            rpt_num_id_sindigito = case_when(
              nchar(rpt_num_id) > 8 ~ substr(rpt_num_id, 1, 9),
              TRUE ~ rpt_num_id
            )
          )
        
        Fondo_Inversion <- Fondo_Inversion %>%
          mutate(
            rpt_valor_absoluto = round(rpt_valor_absoluto, 0)
          )
        
        VIVOS_G3_CDT_fondo_inversion = left_join(VIVOS_G3_CDT_fondo_inversion, Fondo_Inversion, by = c("rpt_num_id", "rpt_valor_absoluto","rpt_fecha_aplicacion"))
        
        #Separamos los que ya cruzaron de los que no
        muertos_1_fondo_inversion <- VIVOS_G3_CDT_fondo_inversion %>%
          filter(!is.na(tipo_fondo) | !is.na(tipo_transaccion))
        
        vivos_1_fondo_inversion <- VIVOS_G3_CDT_fondo_inversion %>%
          filter(is.na(tipo_fondo) & is.na(tipo_transaccion))       
        
        
        #Segunda validacion sin digito de verificacion
        
        #Quitamos de la tabla de inversión todo lo que ya cruzo
        Fondo_Inversion_limpio <- Fondo_Inversion %>%
          anti_join(muertos_1_fondo_inversion, by = c("rpt_num_id", "rpt_valor_absoluto","rpt_fecha_aplicacion"))
        
        #Cambiamos el nombre de la columna de NIT
        names(Fondo_Inversion_limpio)= c("rpt_num_id_sindigito","tipo_fondo","tipo_transaccion","rpt_valor_absoluto","rpt_fecha_aplicacion")
        
        #Quitamos las ultimas dos columas de las vivas
        vivos_1_fondo_inversion <- vivos_1_fondo_inversion %>% select(-tipo_fondo, -tipo_transaccion)
        
        #Volvemos a cruzar los limpios
        VIVOS_G3_CDT_fondo_inversion_2 = left_join(vivos_1_fondo_inversion, Fondo_Inversion_limpio, by = c("rpt_num_id_sindigito", "rpt_valor_absoluto","rpt_fecha_aplicacion"))
        
        muertos_2_fondo_inversion <- VIVOS_G3_CDT_fondo_inversion_2 %>%
          filter(!is.na(tipo_fondo) | !is.na(tipo_transaccion))
        
        vivos_2_fondo_inversion <- VIVOS_G3_CDT_fondo_inversion_2 %>%
          filter(is.na(tipo_fondo) & is.na(tipo_transaccion))  
        
        #Concatenamos los muertos
        muertos_fondos_inversion = rbind(muertos_1_fondo_inversion,muertos_2_fondo_inversion)
        
        #Volvemos todos los datos string sin caracteres UFT8 para poder sacar la tabla
        muertos_fondos_inversion <- muertos_fondos_inversion %>%
          mutate(across(where(is.character), ~ stri_enc_toutf8(., validate = TRUE)))
        
        #Sacamos la tabla para la validacion
        index <-i
        var_name_fondos <- paste("Cruza_fondos_14903_", index, sep = "")
        assign(var_name_fondos, muertos_fondos_inversion, envir = .GlobalEnv)
        
        
        #Sacamos Los que no cruzaron los vivos para el siguente cruce
        VIVOS_G3_CDT <- VIVOS_G3_CDT %>%
          anti_join(muertos_fondos_inversion, by = c("rpt_num_id", "rpt_valor","rpt_fecha_aplicacion"))
      }, 
      error = function(e){ 
        mensaje = paste('El archivo', fondo_de_inversion, 'no esta disponible')
        cat(mensaje, "\n")
      }
      )
      
      
      #============================= Tablas dinamicas de los datos vivos finales G3 -14903 ===============
      Resumen_cancelacion_plazo_FIJO = VIVOS_G3_CDT %>% mutate(naturaleza = ifelse(rpt_tip_id == "03" | rpt_tip_id == 3,"Juridico", "Natural" ))
      
      Resumen_cancelacion_plazo_FIJO_2 = Resumen_cancelacion_plazo_FIJO %>% 
        group_by(en_nomlar2, naturaleza) %>% 
        summarise(rpt_valor = sum(rpt_valor))%>%
        pivot_wider(names_from = naturaleza, values_from = rpt_valor) 
      
      Resumen_cancelacion_plazo_FIJO_2 = Resumen_cancelacion_plazo_FIJO_2 %>% mutate(Juridico = ifelse(is.na(Juridico), 0, Juridico),
                                                                                     Natural = ifelse(is.na(Natural), 0, Natural),
                                                                                     Total = Juridico + Natural)
      Total_resumen = data.frame(en_nomlar2 = "Total",
                                 Juridico = sum(Resumen_cancelacion_plazo_FIJO_2$Juridico),
                                 Natural = sum(Resumen_cancelacion_plazo_FIJO_2$Natural) ,
                                 Total = sum(Resumen_cancelacion_plazo_FIJO_2$Total))
      
      Resumen_cancelacion_plazo_FIJO_2 = rbind(Resumen_cancelacion_plazo_FIJO_2,Total_resumen)
      
      Resumen_cancelacion_plazo_FIJO_2_EFEC = Resumen_cancelacion_plazo_FIJO_2 %>% filter(en_nomlar2 =="EFEC")
      Resumen_cancelacion_plazo_FIJO_2_Total = Resumen_cancelacion_plazo_FIJO_2 %>% filter(en_nomlar2 =="Total")
      
      #============================= Contribución asumida ================================
      #14901 Y 14903
      Vivos_G2_contribucion = VIVOS_G2_CDT_4 %>% mutate(naturaleza = ifelse(rpt_tip_id == "03" | rpt_tip_id == 3,"Juridico", "Natural" )) 
      tabla_dinamica_G2_contribucion = Vivos_G2_contribucion %>% group_by(rpt_num_id, en_nomlar2, naturaleza) %>% summarise(apertura = sum(rpt_valor))
      
      Vivos_G3_contribucion = VIVOS_G3_CDT %>% mutate(naturaleza = ifelse(rpt_tip_id == "03" | rpt_tip_id == 3,"Juridico", "Natural" ))
      tabla_dinamica_G3_contribucion = Vivos_G3_contribucion %>%  group_by(rpt_num_id,en_nomlar2, naturaleza) %>% summarise(Cancelacion_plazo_fijo = sum(rpt_valor))
      
      
      #Juntamos las doa agrupaciones anteriores para poder sumar sus valores totalizados
      CRUCE_5_CDT = full_join(tabla_dinamica_G2_contribucion, tabla_dinamica_G3_contribucion, by= c("rpt_num_id", "en_nomlar2", "naturaleza"))
      
      
      CRUCE_5_CDT$apertura = ifelse(is.na(CRUCE_5_CDT$apertura),0,CRUCE_5_CDT$apertura)
      CRUCE_5_CDT$Cancelacion_plazo_fijo = ifelse(is.na(CRUCE_5_CDT$Cancelacion_plazo_fijo),0,CRUCE_5_CDT$Cancelacion_plazo_fijo)
      CRUCE_5_CDT = CRUCE_5_CDT %>% mutate(control =apertura +  Cancelacion_plazo_fijo)
      CONTRIBUCION_CRUCE = CRUCE_5_CDT %>% mutate(GMF = ifelse(control<0,control*4/1000,0))
      CRUCE_5_CDT = CRUCE_5_CDT[order(CRUCE_5_CDT$control),]
#Separación por especie y naturaleza
      Positivos_Efectivos = CONTRIBUCION_CRUCE %>% filter(control > 0 & en_nomlar2 == "EFEC")
      Positivos_Otros = CONTRIBUCION_CRUCE %>% filter(control > 0 & en_nomlar2 != "EFEC")
      
      Negativos_Efectivos = CONTRIBUCION_CRUCE %>% filter(control <= 0 & en_nomlar2 == "EFEC")
      Negativos_Otros = CONTRIBUCION_CRUCE %>% filter(control <= 0 & en_nomlar2 != "EFEC")
      
      #================================= Presentacion 
      Resumen_contribucion_efectivo_positivo = Positivos_Efectivos %>% 
        group_by(naturaleza) %>% 
        summarise(apertura = sum(apertura),
                  Cancelacion_plazo_fijo = sum(Cancelacion_plazo_fijo),
                  control = sum(control),
                  GMF = sum(GMF)) %>% mutate(GMF_Cancelacion = (Cancelacion_plazo_fijo * 4/1000)*-1)
      
      Resumen_contribucion_efectivo_positivo_Natural = filter(Resumen_contribucion_efectivo_positivo, Resumen_contribucion_efectivo_positivo$naturaleza == "Natural")
      Resumen_contribucion_efectivo_positivo_Juridica = filter(Resumen_contribucion_efectivo_positivo, Resumen_contribucion_efectivo_positivo$naturaleza == "Juridico")
      
      Total_positivos_otros = sum(Positivos_Otros$Cancelacion_plazo_fijo)
      Gmf_positivos_otros = (Total_positivos_otros * 4/1000)*-1
      
      Resumen_contribucion_efectivo_negativo = Negativos_Efectivos %>% 
        group_by(naturaleza) %>% 
        summarise(apertura = sum(apertura),
                  Cancelacion_plazo_fijo = sum(Cancelacion_plazo_fijo),
                  control = sum(control),
                  GMF = sum(GMF)) %>% mutate(GMF_Apertura = apertura * 4/1000 )
      
      Resumen_contribucion_efectivo_negativo_Natural = filter(Resumen_contribucion_efectivo_negativo, Resumen_contribucion_efectivo_negativo$naturaleza == "Natural")
      Resumen_contribucion_efectivo_negativo_Juridico = filter(Resumen_contribucion_efectivo_negativo, Resumen_contribucion_efectivo_negativo$naturaleza == "Juridico")
      
      Total_negativos_otros = sum(Negativos_Otros$apertura)
      Gmf_negativos_otros = Total_negativos_otros * 4/1000
      
      resumen_contribucion = data.frame(Concepto = c("Natural_EFEC", "Juridico_EFEC", "Otros"),
                                        Apertura_14901 = c(Resumen_contribucion_efectivo_negativo_Natural$apertura, Resumen_contribucion_efectivo_negativo_Juridico$apertura , Total_negativos_otros),
                                        GMF_14901 = c(Resumen_contribucion_efectivo_negativo_Natural$GMF_Apertura, Resumen_contribucion_efectivo_negativo_Juridico$GMF_Apertura , Gmf_negativos_otros),
                                        Cancelacion_14903 = c(Resumen_contribucion_efectivo_positivo_Natural$Cancelacion_plazo_fijo, Resumen_contribucion_efectivo_positivo_Juridica$Cancelacion_plazo_fijo ,Total_positivos_otros),
                                        GMF_14903 = c(Resumen_contribucion_efectivo_positivo_Natural$GMF_Cancelacion,Resumen_contribucion_efectivo_positivo_Juridica$GMF_Cancelacion,Gmf_positivos_otros)) %>% mutate(Contribucion = GMF_14901 + GMF_14903)
      
      resumen_contribucion_natural =  resumen_contribucion %>% filter(Concepto == "Natural_EFEC")
      resumen_contribucion_juridico = resumen_contribucion %>% filter(Concepto == "Juridico_EFEC")
      resumen_contribucion_otros = resumen_contribucion %>% filter(Concepto == "Otros")
      
      
      #================================= Formato semanal a reportar ======================
      formato_semal = data.frame(concepto = c("Natural-EFEC",                                 "Juridica-EFEC",                                  "Otros",                                                                                    "Otros-Redenciones"),
                                 Base = c(Resumen_cancelacion_plazo_FIJO_2_EFEC$Natural,      Resumen_cancelacion_plazo_FIJO_2_EFEC$Juridico,   Resumen_cancelacion_plazo_FIJO_2_Total$Total - Resumen_cancelacion_plazo_FIJO_2_EFEC$Total, Resumen_cancelacion_plazo_FIJO_2_Total$Total - Resumen_cancelacion_plazo_FIJO_2_EFEC$Total) ,
                                 Contribucion = c(resumen_contribucion_natural$Contribucion,  resumen_contribucion_juridico$Contribucion,       resumen_contribucion_otros$Contribucion,                                                    0) )
      
      formato_semal = formato_semal %>% mutate(GMF = Base*4/1000, 
                                               Total = GMF - formato_semal$Contribucion)
      formato_semal = formato_semal[,c("concepto","Base","GMF","Contribucion","Total")]
      
      formato_semal = formato_semal %>% mutate(Base = ifelse(Base <0, Base*-1, Base),
                                               GMF = ifelse(GMF <0, GMF*-1, GMF),
                                               Contribucion = ifelse(Contribucion <0, Contribucion*-1, Contribucion),
                                               Total = GMF - Contribucion  )
      
      #================================= Intereses certificados ==========================
      intereses_certificados = data.frame(Concepto = c("Intereses", "Intereses CDTS", "Intereses -Deceval"),
                                          Base = c(total_14943_Pago_Intereses_otros + total_cdt_deceval, total_14943_Pago_Intereses_otros,total_cdt_deceval))
      intereses_certificados = intereses_certificados %>% mutate(GMF = Base * 4/1000,
                                                                 Contribucion = 0,
                                                                 Total = GMF)
      intereses_certificados = intereses_certificados %>% mutate(Base = ifelse(Base <0, Base*-1, Base),
                                                                 GMF = ifelse(GMF <0, GMF*-1, GMF),
                                                                 Contribucion = 0,
                                                                 Total = ifelse(Total <0, Total*-1, Total))
      
      # ========================== Generar un nombre de variable con un índice =====================
      index <-i
      var_name_1 <- paste("DECEVAL_DAVIVIENDA_", index, sep = "")
      var_name_2 <- paste("Anulacion_VS_Apertura_", index, sep = "")
      var_name_3 <- paste("Intereses_", index, sep = "")
      var_name_4 <- paste("tabla_dinamica_14919_vs_14903_", index, sep = "")
      var_name_5 <- paste("14919_VS_14903_", index, sep = "")
      var_name_8 <- paste("tabla_dinamica_14901vs14919_", index, sep = "")
      var_name_9 <- paste("14901_VS_14919_", index, sep = "")
      var_name_10 <- paste("Resumen_14903_", index, sep = "")
      var_name_11 <- paste("tabla_dinamica_contribucion_", index, sep = "")
      var_name_12 <- paste("Contribucion_", index, sep = "")
      var_name_14 <- paste("Formato_semanal_", index, sep = "")
      var_name_15 <- paste("Intereses_Certificados_", index, sep = "")
      var_name_16 <- paste("Nombre_R_CDT_", index, sep = "")
      var_name_17 <- paste("Nombre_R_CDT_Conceptos_", index, sep = "")
      #var_name_18 <- paste("Cruce_fondo_inversion_", index, sep = "")
      
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, DECEVAL_DAVIVIENDA, envir = .GlobalEnv)
      assign(var_name_2, muertos_final_completo_CRUCE_1, envir = .GlobalEnv)
      assign(var_name_3, Intereses, envir = .GlobalEnv)
      assign(var_name_4, Tabla_dinamica_full, envir = .GlobalEnv)
      assign(var_name_5, Muertos_completo_G3_VS_G4, envir = .GlobalEnv)
      assign(var_name_8, Tabla_dinamica_full_tercera_parte, envir = .GlobalEnv)
      assign(var_name_9, Muertos_completo_G2_VS_G4_4, envir = .GlobalEnv)
      assign(var_name_10, Resumen_cancelacion_plazo_FIJO_2, envir = .GlobalEnv)
      assign(var_name_11, CRUCE_5_CDT, envir = .GlobalEnv)
      assign(var_name_12, resumen_contribucion, envir = .GlobalEnv)
      assign(var_name_14, formato_semal, envir = .GlobalEnv)
      assign(var_name_15, intereses_certificados, envir = .GlobalEnv)
      assign(var_name_16, var_name_14, envir = .GlobalEnv)
      assign(var_name_17, var_name_15, envir = .GlobalEnv)
      
      
      #if (exists(fondo_de_inversion)) {
      #  assign(var_name_18, muertos_fondos_inversion, envir = .GlobalEnv)
      #} else {
      #  print("No esta cargado el archivo de fonde de inversiones ")
      #}
      
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}
#cambio

CDT_Semana_1 = "CDT_Semana_1.csv"
CDT_Semana_2 = "CDT_Semana_2.csv"
CDT_Semana_3 = "CDT_Semana_3.csv"
CDT_Semana_4 = "CDT_Semana_4.csv"
CDT_Semana_5 = "CDT_Semana_5.csv"
CDT_Completo = "CDT_Completo.csv"

Fondo_Inversion_Semana_1 = "Fondo_Inversion_Semana_1.csv"
Fondo_Inversion_Semana_2 = "Fondo_Inversion_Semana_2.csv"
Fondo_Inversion_Semana_3 = "Fondo_Inversion_Semana_3.csv"
Fondo_Inversion_Semana_4 = "Fondo_Inversion_Semana_4.csv"
Fondo_Inversion_Semana_5 = "Fondo_Inversion_Semana_5.csv"
Fondo_Inversion_Completo = "Fondo_Inversion_Completo.csv"

Redenciones_CDTS_Semana_1 = "Redenciones_CDTS_Semana_1.DAT"
Redenciones_CDTS_Semana_2 = "Redenciones_CDTS_Semana_2.DAT"
Redenciones_CDTS_Semana_3 = "Redenciones_CDTS_Semana_3.DAT"
Redenciones_CDTS_Semana_4 = "Redenciones_CDTS_Semana_4.DAT"
Redenciones_CDTS_Semana_5 = "Redenciones_CDTS_Semana_5.DAT"
Redenciones_CDTS_Completo = "Redenciones_CDTS_Completo.DAT"

CDT(CDT_Semana_1, 1,Fondo_Inversion_Semana_1,Redenciones_CDTS_Semana_1) 
CDT(CDT_Semana_2, 2,Fondo_Inversion_Semana_2,Redenciones_CDTS_Semana_2) 
CDT(CDT_Semana_3, 3,Fondo_Inversion_Semana_3,Redenciones_CDTS_Semana_3) 
CDT(CDT_Semana_4, 4,Fondo_Inversion_Semana_4,Redenciones_CDTS_Semana_4) 
CDT(CDT_Semana_5, 5,Fondo_Inversion_Semana_5,Redenciones_CDTS_Semana_5) 
CDT(CDT_Completo, 6,Fondo_Inversion_Semana_6,Redenciones_CDTS_Semana_6) 

# .TXT
Redenciones_CDTS_Semana_1_TXT = "Redenciones_CDTS_Semana_1.TXT"
Redenciones_CDTS_Semana_2_TXT = "Redenciones_CDTS_Semana_2.TXT"
Redenciones_CDTS_Semana_3_TXT = "Redenciones_CDTS_Semana_3.TXT"
Redenciones_CDTS_Semana_4_TXT = "Redenciones_CDTS_Semana_4.TXT"
Redenciones_CDTS_Semana_5_TXT = "Redenciones_CDTS_Semana_5.TXT"
Redenciones_CDTS_Completo_TXT = "Redenciones_CDTS_Completo.TXT"

CDT(CDT_Semana_1, 1,Fondo_Inversion_Semana_1,Redenciones_CDTS_Semana_1_TXT) 
CDT(CDT_Semana_2, 2,Fondo_Inversion_Semana_2,Redenciones_CDTS_Semana_2_TXT) 
CDT(CDT_Semana_3, 3,Fondo_Inversion_Semana_3,Redenciones_CDTS_Semana_3_TXT) 
CDT(CDT_Semana_4, 4,Fondo_Inversion_Semana_4,Redenciones_CDTS_Semana_4_TXT) 
CDT(CDT_Semana_5, 5,Fondo_Inversion_Semana_5,Redenciones_CDTS_Semana_5_TXT) 
CDT(CDT_Completo, 6,Fondo_Inversion_Semana_6,Redenciones_CDTS_Completo_TXT)

#================================= Cheques girados (CH gastos) ===============================
cheques_girados = function(datos, i, parametros_cheques_girados) {
  tryCatch(
    {
      datos = read_excel (datos, skip = 11)
      names(datos) = c("UN", "G_LIBROS","LIBRO","CUENTA", "SUCURSAL","DEPENDENCIA","ID_DE_ASIENTO","FECHA_COMPROBANTE","FECHA_PROCESO","DESCRIPCION","DEBITO","CREDITO","AUXILIAR","REFERENCIA","USUARIO","ID_COMPROBANTE" , "ESTADO","REAL")  
      datos = datos[,c("DEBITO","CREDITO", "ID_COMPROBANTE", "ID_DE_ASIENTO")]
      datos = datos %>% mutate( across(c(1:2), ~ as.numeric(.)),
                                 ID_COMPROBANTE = substr(ID_COMPROBANTE, 1,2),
                                 ID_DE_ASIENTO = substr(ID_DE_ASIENTO, 1,2))
      datos = data.frame(datos)
      
      #Hacemos un join para traer los datos 
      datos = left_join(datos, parametros_cheques_girados, by = c("ID_COMPROBANTE", "ID_DE_ASIENTO"))
      datos = data.frame(datos)
      
      #Aqui hacemos un control para cheques por si resulta alguna otra combinacion 
      datos_sif  = filter(datos, datos$Concepto == "Sif")
      datos_cheques_girados = filter(datos, datos$Concepto == "Cheques Girados") 
      
      Conceptos = c("Cheques girados", "Sif", "Total_por_conceptos", "Total_Base")
      totales_debito = c(abs(sum(datos_cheques_girados$DEBITO)), abs(sum(datos_sif$DEBITO)), abs(sum(datos_cheques_girados$DEBITO) + sum(datos_sif$DEBITO)), abs(sum(datos$DEBITO)))
      totales_credito = c(abs(sum(datos_cheques_girados$CREDITO)), abs(sum(datos_sif$CREDITO)), abs(sum(datos_cheques_girados$CREDITO) + sum(datos_sif$CREDITO)), abs(sum(datos$CREDITO)))
      Control_cheques_girados = data.frame(Concepto = Conceptos , Debito = totales_debito , Credito = totales_credito)
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("cheques_girados_", index, sep = "")
      var_name_2 <- paste("R_cheques_girados_", index, sep = "")
      var_name_3 <- paste("Nombre_cheques_girados_", index, sep = "")
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, datos, envir = .GlobalEnv)
      assign(var_name_2, Control_cheques_girados, envir = .GlobalEnv)
      assign(var_name_3, var_name_2, envir = .GlobalEnv)
      
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}

Cheques_Girados_Semana_1 = "Cheques_Girados_Semana_1.xlsx"
Cheques_Girados_Semana_2 = "Cheques_Girados_Semana_2.xlsx"
Cheques_Girados_Semana_3 = "Cheques_Girados_Semana_3.xlsx"
Cheques_Girados_Semana_4 = "Cheques_Girados_Semana_4.xlsx"
Cheques_Girados_Semana_5 = "Cheques_Girados_Semana_5.xlsx"
Cheques_Girados_Completo = "Cheques_Girados_Completo.xlsx"

cheques_girados(Cheques_Girados_Semana_1, 1, parametros_cheques_girados) 
cheques_girados(Cheques_Girados_Semana_2, 2, parametros_cheques_girados) 
cheques_girados(Cheques_Girados_Semana_3, 3, parametros_cheques_girados) 
cheques_girados(Cheques_Girados_Semana_4, 4, parametros_cheques_girados) 
cheques_girados(Cheques_Girados_Semana_5, 5, parametros_cheques_girados) 
cheques_girados(Cheques_Girados_Completo, 6, parametros_cheques_girados) 



#================================= Pago de Intereses ===============================
Pago_Intereses = function(datos, i) {
  tryCatch(
    {
      datos <- read_excel(datos, col_types = c("text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "numeric", "text", "text", "text", "text", "text", "text", "numeric", "numeric", "numeric", "text"), skip = 9)
      names(datos) = c("Cuenta", "Descripcion_Cuenta","Sucursal","Descripcion_Sucursal","Oficina",  "Descripcion_Oficina","Tipo_Comprobante","Descripcion_Tipo_Comprobante","Numero_de_Transaccion","Descripcion_Transaccion","Fecha_Transaccion","Fecha_de_grabacion","Descripcion_Cabecera", "Numero_de_Identificacion","Digito_de_Verificacion","Razon_Social","Documento","Referencia_1","Dependencia","Subproyecto","Auxiliar_de_Conciliacion", "Tipo_de_Evento","Clase_Contable", "Numero_de_Linea","Saldo_Dia_Anterior","Debito","Credito","Asiento_Reversado")
      datos = datos[,c("Cuenta","Debito","Credito")]
      datos = datos %>% mutate(across(c(2:3), ~ as.numeric(.)),
                               Debito  = ifelse(is.na(Debito), 0, Debito),
                               Credito = ifelse(is.na(Credito), 0, Credito),
                               Neto = Debito - Credito)
      datos = data.frame(datos)
      
      #Hacemos una agrupación por cuentas 
      datos = data.frame(datos  %>% group_by(Cuenta) %>% summarise(Debito = sum(Debito), Credito = sum(Credito), Neto =sum(Neto))) %>% mutate(GMF = Neto * 0.004)
      datos = filter(datos, datos$Cuenta == "5102050010" | datos$Cuenta == "5102050028" | datos$Cuenta == "5102950011" )
      
      C5102050010 = filter(datos, datos$Cuenta == "5102050010")
      C5102050028 = filter(datos, datos$Cuenta == "5102050028")
      C5102950011 = filter(datos, datos$Cuenta == "5102950011")
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("Pago_Intereses_", index, sep = "")
      var_name_2 <- paste("C5102050010_", index, sep = "")
      var_name_3 <- paste("C5102050028_", index, sep = "")
      var_name_4 <- paste("C5102950011_", index, sep = "")
      var_name_5 <- paste("Nombre_Damas_", index, sep = "")
      var_name_6 <- paste("Nombre_Fijo_Diario_", index, sep = "")
      var_name_7 <- paste("Nombre_Cuenta_Corriente_", index, sep = "")
      
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, datos, envir = .GlobalEnv)
      assign(var_name_2, C5102050010, envir = .GlobalEnv)
      assign(var_name_3, C5102050028, envir = .GlobalEnv)
      assign(var_name_4, C5102950011, envir = .GlobalEnv)
      assign(var_name_5, var_name_3, envir = .GlobalEnv)
      assign(var_name_6, var_name_2, envir = .GlobalEnv)
      assign(var_name_7, var_name_4, envir = .GlobalEnv)

      
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}

Pago_Intereses_Semana_1 = "Pago_Intereses_Semana_1.xlsx"
Pago_Intereses_Semana_2 = "Pago_Intereses_Semana_2.xlsx"
Pago_Intereses_Semana_3 = "Pago_Intereses_Semana_3.xlsx"
Pago_Intereses_Semana_4 = "Pago_Intereses_Semana_4.xlsx"
Pago_Intereses_Semana_5 = "Pago_Intereses_Semana_5.xlsx"
Pago_Intereses_Completo = "Pago_Intereses_Completo.xlsx"

Pago_Intereses(Pago_Intereses_Semana_1, 1) 
Pago_Intereses(Pago_Intereses_Semana_2, 2) 
Pago_Intereses(Pago_Intereses_Semana_3, 3) 
Pago_Intereses(Pago_Intereses_Semana_4, 4) 
Pago_Intereses(Pago_Intereses_Semana_5, 5) 
Pago_Intereses(Pago_Intereses_Completo, 6) 





#================================= Timbre ===============================

Timbre = function(datos, i) {
  tryCatch(
    {
      datos <- read_excel(datos, sheet = "Timbre_0098", col_types = c("text", "text", "date", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text", "text"))
      names(datos) = c("COMPROBANTE","CODOFIC","FECHA", "DESCRIPCION","DOCUMENTO","VALOR","MOTIVO","Otro_Cual", "No_CHEQUE_ANULADO_No_DE_CREDITO_GIRADO","No_CHEQUE_REPOSICION","No_CHEQUERA","BENEFICIARIO_TITULAR","NIT","GMF_ASUMIDO", "semana")
      datos = datos[,c("VALOR","GMF_ASUMIDO")]
      datos = datos %>% mutate(VALOR = str_replace_all(VALOR,"[,]", ""),
                               VALOR  = as.numeric(VALOR),
                               GMF_ASUMIDO = toupper(GMF_ASUMIDO),
                               GMF_ASUMIDO = str_remove_all(GMF_ASUMIDO, " "))
      datos = data.frame(datos)
      
      #Filtramos por lo que digan que SI
      datos = filter(datos, datos$GMF_ASUMIDO == "SI")
      Contribucion_Asumida = sum(datos$VALOR) * 0.004
      
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("Contribucion_Asumida_", index, sep = "")
      var_name_2 <- paste("Nombre_Contribucion_Asumida_", index, sep = "")
      

      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, Contribucion_Asumida, envir = .GlobalEnv)
      assign(var_name_2, var_name_1, envir = .GlobalEnv)
      
    },
    error = function(e){ 
      mensaje = paste('El archivo', datos, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}

Timbre_Semana_1 = "Timbre_Semana_1.xlsx"
Timbre_Semana_2 = "Timbre_Semana_2.xlsx"
Timbre_Semana_3 = "Timbre_Semana_3.xlsx"
Timbre_Semana_4 = "Timbre_Semana_4.xlsx"
Timbre_Semana_5 = "Timbre_Semana_5.xlsx"
Timbre_Completo = "Timbre_Completo.xlsx"

Timbre(Timbre_Semana_1, 1) 
Timbre(Timbre_Semana_2, 2) 
Timbre(Timbre_Semana_3, 3) 
Timbre(Timbre_Semana_4, 4) 
Timbre(Timbre_Semana_5, 5) 
Timbre(Timbre_Completo, 6)
#================================= Datos en el comprobante ===============================
#Cheques
R_cheques_funcion = function (data, nombre_data, fila1, fila2){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,4]
    comprobante_1[fila1,8] <<- data[1,5]
    comprobante_1[fila2,7] <<- data[5,4]
    comprobante_1[fila2,8] <<- data[5,5]
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("R_cheques_oficina_1")){R_cheques_funcion(R_cheques_oficina_1, Nombre_R_cheques_1,  1, 65)} else {print("No esta disponible 1")}
if (exists("R_cheques_oficina_2")){R_cheques_funcion(R_cheques_oficina_2, Nombre_R_cheques_2,  2, 66)} else {print("No esta disponible 2")}
if (exists("R_cheques_oficina_3")){R_cheques_funcion(R_cheques_oficina_3, Nombre_R_cheques_3,  3, 67)} else {print("No esta disponible 3")}
if (exists("R_cheques_oficina_4")){R_cheques_funcion(R_cheques_oficina_4, Nombre_R_cheques_4,  4, 68)} else {print("No esta disponible 4")}
if (exists("R_cheques_oficina_5")){R_cheques_funcion(R_cheques_oficina_5, Nombre_R_cheques_5,  5, 69)} else {print("No esta disponible 5")}
if (exists("R_cheques_oficina_6")){R_cheques_funcion(R_cheques_oficina_6, Nombre_R_cheques_6,  7, 71)} else {print("No esta disponible 6")}

#ISA
R_ISA_funcion = function (data, nombre_data, fila1, fila2){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,2]
    comprobante_1[fila1,8] <<- data[1,3]
    comprobante_1[fila2,7] <<- data[2,2]
    comprobante_1[fila2,8] <<- data[2,3]
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("R_ISA_sobrantes_1")){R_ISA_funcion(R_ISA_sobrantes_1, Nombre_R_ISA_1, 81 , 73)} else {print("No esta disponible 1")}
if (exists("R_ISA_sobrantes_2")){R_ISA_funcion(R_ISA_sobrantes_2, Nombre_R_ISA_2, 82 , 74)} else {print("No esta disponible 2")}
if (exists("R_ISA_sobrantes_3")){R_ISA_funcion(R_ISA_sobrantes_3, Nombre_R_ISA_3, 83 , 75)} else {print("No esta disponible 3")}
if (exists("R_ISA_sobrantes_4")){R_ISA_funcion(R_ISA_sobrantes_4, Nombre_R_ISA_4, 84 , 76)} else {print("No esta disponible 4")}
if (exists("R_ISA_sobrantes_5")){R_ISA_funcion(R_ISA_sobrantes_5, Nombre_R_ISA_5, 85 , 77)} else {print("No esta disponible 5")}
if (exists("R_ISA_sobrantes_6")){R_ISA_funcion(R_ISA_sobrantes_6, Nombre_R_ISA_6, 87 , 79)} else {print("No esta disponible 6")}

#Provedores
R_Provedores_funcion = function (data, nombre_data, fila1, fila2, fila3){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,5]
    comprobante_1[fila1,8] <<- data[1,6]
    comprobante_1[fila2,7] <<- data[2,5]
    comprobante_1[fila2,8] <<- data[2,6]
    comprobante_1[fila3,7] <<- data[3,5]
    comprobante_1[fila3,8] <<- data[3,6]
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("R_pago_proveedores_1")){R_Provedores_funcion(R_pago_proveedores_1, Nombre_R_Provedores_1, 113 , 121, 129)} else {print("No esta disponible 1")}
if (exists("R_pago_proveedores_2")){R_Provedores_funcion(R_pago_proveedores_2, Nombre_R_Provedores_2, 114 , 122, 130)} else {print("No esta disponible 2")}
if (exists("R_pago_proveedores_3")){R_Provedores_funcion(R_pago_proveedores_3, Nombre_R_Provedores_3, 115 , 123, 131)} else {print("No esta disponible 3")}
if (exists("R_pago_proveedores_4")){R_Provedores_funcion(R_pago_proveedores_4, Nombre_R_Provedores_4, 116 , 124, 132)} else {print("No esta disponible 4")}
if (exists("R_pago_proveedores_5")){R_Provedores_funcion(R_pago_proveedores_5, Nombre_R_Provedores_5, 117 , 125, 133)} else {print("No esta disponible 5")}
if (exists("R_pago_proveedores_6")){R_Provedores_funcion(R_pago_proveedores_6, Nombre_R_Provedores_6, 119 , 127, 135)} else {print("No esta disponible 6")}

#CDAT
R_CDAT_funcion = function (data, nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,1]
    comprobante_1[fila1,8] <<- data[1,2]
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("resumen_CDAT_1")){R_CDAT_funcion(resumen_CDAT_1, Nombre_R_CDAT_1, 57)} else {print("No esta disponible 1")}
if (exists("resumen_CDAT_2")){R_CDAT_funcion(resumen_CDAT_2, Nombre_R_CDAT_2, 58)} else {print("No esta disponible 2")}
if (exists("resumen_CDAT_3")){R_CDAT_funcion(resumen_CDAT_3, Nombre_R_CDAT_3, 59)} else {print("No esta disponible 3")}
if (exists("resumen_CDAT_4")){R_CDAT_funcion(resumen_CDAT_4, Nombre_R_CDAT_4, 60)} else {print("No esta disponible 4")}
if (exists("resumen_CDAT_5")){R_CDAT_funcion(resumen_CDAT_5, Nombre_R_CDAT_5, 61)} else {print("No esta disponible 5")}
if (exists("resumen_CDAT_6")){R_CDAT_funcion(resumen_CDAT_6, Nombre_R_CDAT_6, 63)} else {print("No esta disponible 6")}

#CDT
R_CDT_resumen_funcion = function (data, nombre_data, fila1, fila2, fila3){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,2]
    comprobante_1[fila1,8] <<- data[1,3]
    comprobante_1[fila1,9] <<- data[1,4]
    comprobante_1[fila1,10] <<- data[1,5]
    
    comprobante_1[fila2,7] <<- data[2,2]
    comprobante_1[fila2,8] <<- data[2,3]
    comprobante_1[fila2,9] <<- data[2,4]
    comprobante_1[fila2,10] <<- data[2,5]
    
    comprobante_1[fila3,7] <<- data[3,2]
    comprobante_1[fila3,8] <<- data[3,3]
    comprobante_1[fila3,9] <<- data[3,4]
    comprobante_1[fila3,10] <<- data[3,5]
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}

if (exists("Formato_semanal_1")){R_CDT_resumen_funcion(Formato_semanal_1, Nombre_R_CDT_1, 17, 25, 33)} else {print("No esta disponible 1")}
if (exists("Formato_semanal_2")){R_CDT_resumen_funcion(Formato_semanal_2, Nombre_R_CDT_2, 18, 26, 34)} else {print("No esta disponible 2")}
if (exists("Formato_semanal_3")){R_CDT_resumen_funcion(Formato_semanal_3, Nombre_R_CDT_3, 19, 27, 35)} else {print("No esta disponible 3")}
if (exists("Formato_semanal_4")){R_CDT_resumen_funcion(Formato_semanal_4, Nombre_R_CDT_4, 20, 28, 36)} else {print("No esta disponible 4")}
if (exists("Formato_semanal_5")){R_CDT_resumen_funcion(Formato_semanal_5, Nombre_R_CDT_5, 21, 29, 37)} else {print("No esta disponible 5")}
if (exists("Formato_semanal_6")){R_CDT_resumen_funcion(Formato_semanal_6, Nombre_R_CDT_6, 23, 31, 39)} else {print("No esta disponible 6")}

#CDT Intereses
R_CDT_conceptos_funcion = function (data, nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,2]
    comprobante_1[fila1,8] <<- data[1,3]
    comprobante_1[fila1,10] <<- data[1,5]
    
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("Intereses_Certificados_1")){R_CDT_conceptos_funcion(Intereses_Certificados_1, Nombre_R_CDT_Conceptos_1, 49)} else {print("No esta disponible 1")}
if (exists("Intereses_Certificados_2")){R_CDT_conceptos_funcion(Intereses_Certificados_2, Nombre_R_CDT_Conceptos_2, 50)} else {print("No esta disponible 2")}
if (exists("Intereses_Certificados_3")){R_CDT_conceptos_funcion(Intereses_Certificados_3, Nombre_R_CDT_Conceptos_3, 51)} else {print("No esta disponible 3")}
if (exists("Intereses_Certificados_4")){R_CDT_conceptos_funcion(Intereses_Certificados_4, Nombre_R_CDT_Conceptos_4, 52)} else {print("No esta disponible 4")}
if (exists("Intereses_Certificados_5")){R_CDT_conceptos_funcion(Intereses_Certificados_5, Nombre_R_CDT_Conceptos_5, 53)} else {print("No esta disponible 5")}
if (exists("Intereses_Certificados_6")){R_CDT_conceptos_funcion(Intereses_Certificados_6, Nombre_R_CDT_Conceptos_6, 55)} else {print("No esta disponible 6")}


#Cheques girados
R_Cheques_Girados_funcion = function (data, nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,3]
    comprobante_1[fila1,8] <<- data[1,3] * 0.004
    comprobante_1[fila1,9] <<- data[1,2] * 0.004
    
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("R_cheques_girados_1")){R_Cheques_Girados_funcion(R_cheques_girados_1, Nombre_cheques_girados_1, 9)} else {print("No esta disponible 1")}
if (exists("R_cheques_girados_2")){R_Cheques_Girados_funcion(R_cheques_girados_2, Nombre_cheques_girados_2, 10)} else {print("No esta disponible 2")}
if (exists("R_cheques_girados_3")){R_Cheques_Girados_funcion(R_cheques_girados_3, Nombre_cheques_girados_3, 11)} else {print("No esta disponible 3")}
if (exists("R_cheques_girados_4")){R_Cheques_Girados_funcion(R_cheques_girados_4, Nombre_cheques_girados_4, 12)} else {print("No esta disponible 4")}
if (exists("R_cheques_girados_5")){R_Cheques_Girados_funcion(R_cheques_girados_5, Nombre_cheques_girados_5, 13)} else {print("No esta disponible 5")}
if (exists("R_cheques_girados_6")){R_Cheques_Girados_funcion(R_cheques_girados_6, Nombre_cheques_girados_6, 15)} else {print("No esta disponible 6")}

#Pago de intereses
R_Pago_Intereses_funcion = function (data, nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,7] <<- data[1,4]
    comprobante_1[fila1,8] <<- data[1,5]
    
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("C5102050028_1")){R_Pago_Intereses_funcion(C5102050028_1, Nombre_Damas_1, 89)} else {print("No esta disponible 1")}
if (exists("C5102050028_2")){R_Pago_Intereses_funcion(C5102050028_2, Nombre_Damas_2, 90)} else {print("No esta disponible 2")}
if (exists("C5102050028_3")){R_Pago_Intereses_funcion(C5102050028_3, Nombre_Damas_3, 91)} else {print("No esta disponible 3")}
if (exists("C5102050028_4")){R_Pago_Intereses_funcion(C5102050028_4, Nombre_Damas_4, 92)} else {print("No esta disponible 4")}
if (exists("C5102050028_5")){R_Pago_Intereses_funcion(C5102050028_5, Nombre_Damas_5, 93)} else {print("No esta disponible 5")}
if (exists("C5102050028_6")){R_Pago_Intereses_funcion(C5102050028_6, Nombre_Damas_6, 95)} else {print("No esta disponible 6")}

if (exists("C5102050010_1")){R_Pago_Intereses_funcion(C5102050010_1, Nombre_Fijo_Diario_1, 97)} else {print("No esta disponible 1")}
if (exists("C5102050010_2")){R_Pago_Intereses_funcion(C5102050010_2, Nombre_Fijo_Diario_2, 98)} else {print("No esta disponible 2")}
if (exists("C5102050010_3")){R_Pago_Intereses_funcion(C5102050010_3, Nombre_Fijo_Diario_3, 99)} else {print("No esta disponible 3")}
if (exists("C5102050010_4")){R_Pago_Intereses_funcion(C5102050010_4, Nombre_Fijo_Diario_4, 100)} else {print("No esta disponible 4")}
if (exists("C5102050010_5")){R_Pago_Intereses_funcion(C5102050010_5, Nombre_Fijo_Diario_5, 101)} else {print("No esta disponible 5")}
if (exists("C5102050010_6")){R_Pago_Intereses_funcion(C5102050010_6, Nombre_Fijo_Diario_6, 103)} else {print("No esta disponible 6")}

if (exists("C5102950011_1")){R_Pago_Intereses_funcion(C5102950011_1, Nombre_Cuenta_Corriente_1, 105)} else {print("No esta disponible 1")}
if (exists("C5102950011_2")){R_Pago_Intereses_funcion(C5102950011_2, Nombre_Cuenta_Corriente_2, 106)} else {print("No esta disponible 2")}
if (exists("C5102950011_3")){R_Pago_Intereses_funcion(C5102950011_3, Nombre_Cuenta_Corriente_3, 107)} else {print("No esta disponible 3")}
if (exists("C5102950011_4")){R_Pago_Intereses_funcion(C5102950011_4, Nombre_Cuenta_Corriente_4, 108)} else {print("No esta disponible 4")}
if (exists("C5102950011_5")){R_Pago_Intereses_funcion(C5102950011_5, Nombre_Cuenta_Corriente_5, 109)} else {print("No esta disponible 5")}
if (exists("C5102950011_6")){R_Pago_Intereses_funcion(C5102950011_6, Nombre_Cuenta_Corriente_6, 111)} else {print("No esta disponible 6")}
#Pago de intereses
R_Contribucion_Asumida_funcion = function (data, nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,9] <<- data
    
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}
if (exists("Contribucion_Asumida_1")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_1, Nombre_Contribucion_Asumida_1, 1)} else {print("No esta disponible 1")}
if (exists("Contribucion_Asumida_2")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_2, Nombre_Contribucion_Asumida_2, 2)} else {print("No esta disponible 2")}
if (exists("Contribucion_Asumida_3")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_3, Nombre_Contribucion_Asumida_3, 3)} else {print("No esta disponible 3")}
if (exists("Contribucion_Asumida_4")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_4, Nombre_Contribucion_Asumida_4, 4)} else {print("No esta disponible 4")}
if (exists("Contribucion_Asumida_5")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_5, Nombre_Contribucion_Asumida_5, 5)} else {print("No esta disponible 5")}
if (exists("Contribucion_Asumida_6")){R_Contribucion_Asumida_funcion(Contribucion_Asumida_6, Nombre_Contribucion_Asumida_6, 7)} else {print("No esta disponible 6")}


#Hacemos que todo lo que este nulo se convierta en cero 
comprobante_1 <- comprobante_1 %>% 
  mutate(across(c(7:11), ~ as.numeric(.)),
         across(c(7:11), ~ ifelse(is.na(.), 0, .)),
         TOTAL_CONTRIBUCION_ME = CALCULO_CONTR - CONTR_ASUMIDA)

#Hacemos el calculo de los controles
control = function(fila) {
  comprobante_1[fila,7] <<- sum(comprobante_1[c((fila - 5): (fila-1)),7])
  comprobante_1[fila,8] <<- sum(comprobante_1[c((fila - 5): (fila-1)),8])
  comprobante_1[fila,9] <<- sum(comprobante_1[c((fila - 5): (fila-1)),9])
  comprobante_1[fila,10] <<- sum(comprobante_1[c((fila - 5): (fila-1)),10])
  comprobante_1[(fila + 2),c(7:10)] <<- comprobante_1[fila, c(7:10)] - comprobante_1[(fila + 1),c(7:10) ] 
  }

control(6)
control(14)
control(22)
control(30)
control(38)
control(46)
control(54)
control(62)
control(70)
control(78)
control(86)
control(94)
control(102)
control(110)
control(118)
control(126)
control(134)


#Hacemos los calculos de la columna que realiza Impuestos 
#Cheques girados + reposiciones
R_Cheques_Girados_Reposiciones_funcion = function (fila1, fila2){ 
   comprobante_1[fila1,11] <<- comprobante_1[fila1,10] + comprobante_1[fila2,10]
   comprobante_1[fila1,11] <<- ifelse(comprobante_1[fila1,11] < 0, 0, comprobante_1[fila1,11])
  }
R_Cheques_Girados_Reposiciones_funcion(1,65)
R_Cheques_Girados_Reposiciones_funcion(2,66)
R_Cheques_Girados_Reposiciones_funcion(3,67)
R_Cheques_Girados_Reposiciones_funcion(4,68)
R_Cheques_Girados_Reposiciones_funcion(5,69)

#Redenciones CDTS - Reposiciones CDTS
R_Redenciones_Reposiciones_funcion = function (nombre_data, fila1){ 
  if (exists(nombre_data)) {
    comprobante_1[fila1,11] <<- comprobante_1[fila1,10]
    
  } else {
    print(paste("Aún no esta disponible", nombre_data))
  }}

if (exists("Nombre_Redenciones_CDTS_Total_1")){R_Redenciones_Reposiciones_funcion(Nombre_Redenciones_CDTS_Total_1, 33)} else {print("No esta disponible 1")}
if (exists("Nombre_Redenciones_CDTS_Total_2")){R_Redenciones_Reposiciones_funcion(Nombre_Redenciones_CDTS_Total_2,  34)} else {print("No esta disponible 2")}
if (exists("Nombre_Redenciones_CDTS_Total_3")){R_Redenciones_Reposiciones_funcion(Nombre_Redenciones_CDTS_Total_3,  35)} else {print("No esta disponible 3")}
if (exists("Nombre_Redenciones_CDTS_Total_4")){R_Redenciones_Reposiciones_funcion(Nombre_Redenciones_CDTS_Total_4,  36)} else {print("No esta disponible 4")}
if (exists("Nombre_Redenciones_CDTS_Total_5")){R_Redenciones_Reposiciones_funcion(Nombre_Redenciones_CDTS_Total_5,  37)} else {print("No esta disponible 5")}


#Vamos a dejar esto por si si aplica y Sandra no se acuerda, se coloca en comentarios
#INTERESES CERT PESOS 
#R_INTERESES_CERT_PESOS_funcion = function (fila1){ 
#  comprobante_1[fila1,11] <<- ifelse(comprobante_1[fila1,10] < 0, 0, comprobante_1[fila1,10])
#}
#R_INTERESES_CERT_PESOS_funcion(49)
#R_INTERESES_CERT_PESOS_funcion(50)
#R_INTERESES_CERT_PESOS_funcion(51)
#R_INTERESES_CERT_PESOS_funcion(52)
#R_INTERESES_CERT_PESOS_funcion(53)


#INTERESES CDAT
#R_INTERESES_CDAT_funcion = function (fila1){ 
#  comprobante_1[fila1,11] <<- comprobante_1[fila1,10]
#}
#R_INTERESES_CDAT_funcion(57)
#R_INTERESES_CDAT_funcion(58)
#R_INTERESES_CDAT_funcion(59)
#R_INTERESES_CDAT_funcion(60)
#R_INTERESES_CDAT_funcion(61)


#Vamos a dejar la columna de impuestos con los totales y sumar semanalmente 
comprobante_1 = comprobante_1 %>%  
  mutate(Columna11 = case_when(TOTAL_CONTRIBUCION_IMPUESTOS == 0 ~ TOTAL_CONTRIBUCION_ME ,
                               TRUE ~ TOTAL_CONTRIBUCION_IMPUESTOS))

comprobante_1$TOTAL_CONTRIBUCION_IMPUESTOS = comprobante_1$Columna11
comprobante_1$Columna11 = NULL

comprobante_1[65:69,11] = 0   

# Crear un vector para almacenar los resultados de las sumas
resultados <- numeric(5)

# Realizar las sumas y guardar los resultados
for (i in 1:5) {
  # Calcular los índices de las filas para la suma actual
  indices <- seq(i, 133, by = 8)
  
  # Sumar los valores de la columna 11 en esos índices
  suma <- sum(comprobante_1[indices, 11])
  
  # Guardar el resultado en el vector de resultados
  resultados[i] <- suma
}

# Asignar los resultados a las filas 137 a 141 de la columna 11
comprobante_1[137:141, 11] <- resultados



#================================= Datos comprobante impuestos ======================
#Eliminar los insumos de la corrida pasada
unlink(local_directory, recursive = TRUE)

#Colocamos el ID del folder de salida
folder_id_salida = as.character(read_sheet("1ZJgei5OBgFTKlv5mydjZetsBoFB9Lc2iCu2E7HVcJwM", sheet = 'Datos_Variables', col_names = FALSE, range = "B17"))

fun_comprobante_impuestos_funcion = function(comprobante_impuestos_funcion, i, Hasta, mes_hasta, numero_semana,  fila_comprobante_me,fila_comprobante_me_1,fila_comprobante_me_2, fila_comprobante_me_3,fila_comprobante_me_4,fila_comprobante_me_5,   fila_comprobante_me_6, fila_comprobante_me_7,fila_comprobante_me_8, fila_comprobante_me_9,fila_comprobante_me_10, fila_comprobante_me_11,fila_comprobante_me_12, fila_comprobante_me_13,fila_comprobante_me_14) {
  tryCatch(
    {
      mes_numero_hasta  = case_when(mes_hasta == "Enero" ~ 1,
                                    mes_hasta == "Febrero" ~ 2, 
                                    mes_hasta == "Marzo" ~ 3,
                                    mes_hasta == "Abril" ~ 4,
                                    mes_hasta == "Mayo" ~ 5,
                                    mes_hasta == "Junio" ~ 6,
                                    mes_hasta == "Julio" ~ 7,
                                    mes_hasta == "Agosto" ~ 8,
                                    mes_hasta == "Septiembre" ~ 9,
                                    mes_hasta == "Octubre" ~ 10,
                                    mes_hasta == "Noviembre" ~ 11,
                                    mes_hasta == "Diciembre" ~ 12, TRUE ~ NA_integer_)
      
      fecha_impuestos = format(as.Date(paste0(Hasta,"/",mes_numero_hasta,"/",año), format = "%d/%m/%Y"), "%d/%m/%Y")
      fecha_impuestos = as.character(fecha_impuestos)
      descripcion_impuestos = paste("CAUSACION PAGO SEMANA", numero_semana)
      
      
      comprobante_impuestos_funcion = data.frame(comprobante_impuestos_funcion) %>% mutate(across(c(18:19), ~ as.numeric(.)),
                                                                                           across(c(18:19), ~ ifelse(is.na(.),0,0)),
                                                                                           FECHA.CONTABLE = fecha_impuestos,
                                                                                           DESCRIPCION = descripcion_impuestos)
      
      #Cheques girados 9710 MOT-40- Consecutivo1:5
      #fila_comprobante_me = 1
      comprobante_impuestos_funcion[9,19] = comprobante_1[fila_comprobante_me, 11]
      comprobante_impuestos_funcion[10,18] = comprobante_1[fila_comprobante_me, 11]
      
      #Cheques girados gastos- Consecutivo 9:13
      #fila_comprobante_me_1 = 9
      comprobante_impuestos_funcion[11,19] = comprobante_1[fila_comprobante_me_1, 11]
      comprobante_impuestos_funcion[12,18] = comprobante_1[fila_comprobante_me_1, 11]
      
      #RED CERTIF-EFECTIVO NATURAL- Consecutivo 17:21
      #fila_comprobante_me_2 = 17
      comprobante_impuestos_funcion[13,19] = comprobante_1[fila_comprobante_me_2, 11]
      comprobante_impuestos_funcion[14,18] = comprobante_1[fila_comprobante_me_2, 11]
      comprobante_impuestos_funcion[15,18] = comprobante_1[fila_comprobante_me_2, 11]
      comprobante_impuestos_funcion[16,19] = comprobante_1[fila_comprobante_me_2, 11]
      
      #RED CERTIF-EFECTIVO JURIDICA- Consecutivo 25:29
      #fila_comprobante_me_3 = 25
      comprobante_impuestos_funcion[17,19] = comprobante_1[fila_comprobante_me_3, 11]
      comprobante_impuestos_funcion[18,18] = comprobante_1[fila_comprobante_me_3, 11]
      comprobante_impuestos_funcion[19,18] = comprobante_1[fila_comprobante_me_3, 11]
      comprobante_impuestos_funcion[20,19] = comprobante_1[fila_comprobante_me_3, 11]
      
      #SOBRANTES FM - Consecutivo 73:77
      #fila_comprobante_me_4 = 73
      comprobante_impuestos_funcion[21,19] = comprobante_1[fila_comprobante_me_4, 11]
      comprobante_impuestos_funcion[22,18] = comprobante_1[fila_comprobante_me_4, 11]
      
      #CHEQUE GIRADO PAGO ISA - Consecutivo:81:85
      #fila_comprobante_me_5 = 81
      comprobante_impuestos_funcion[23,19] = comprobante_1[fila_comprobante_me_5, 11]
      comprobante_impuestos_funcion[24,18] = comprobante_1[fila_comprobante_me_5, 11]
      
      #RED CERTIF OTROS - Consecutivo 33-37
      #fila_comprobante_me_6 = 33
      comprobante_impuestos_funcion[45,19] = comprobante_1[fila_comprobante_me_6, 11]
      comprobante_impuestos_funcion[46,18] = comprobante_1[fila_comprobante_me_6, 11]
      
      #INTERESES CERT CDT -Consecutivo 49-53
      #fila_comprobante_me_7 = 49
      comprobante_impuestos_funcion[51,19] = comprobante_1[fila_comprobante_me_7, 11]
      comprobante_impuestos_funcion[52,18] = comprobante_1[fila_comprobante_me_7, 11]
      
      #INTERESES CERT CDAT - Consecutivo 57:61
      #fila_comprobante_me_8 = 57
      comprobante_impuestos_funcion[53,19] = comprobante_1[fila_comprobante_me_8, 11]
      comprobante_impuestos_funcion[54,18] = comprobante_1[fila_comprobante_me_8, 11]
      
      #PAGO PROVEEDORES DAMAS - Consecutivo 113:117
      #fila_comprobante_me_9 = 113
      comprobante_impuestos_funcion[73,19] = comprobante_1[fila_comprobante_me_9, 11]
      comprobante_impuestos_funcion[74,18] = comprobante_1[fila_comprobante_me_9, 11]
      
      #PAGO PROVEEDORES FIJO DIARIO - Consecutivo  121:125
      #fila_comprobante_me_10 = 121
      comprobante_impuestos_funcion[75,19] = comprobante_1[fila_comprobante_me_10, 11]
      comprobante_impuestos_funcion[76,18] = comprobante_1[fila_comprobante_me_10, 11]
      
      #PAGO PROVEEDORES CTA CTE - Consecutivo 129:133
      #fila_comprobante_me_11 = 129
      comprobante_impuestos_funcion[77,19] = comprobante_1[fila_comprobante_me_11, 11]
      comprobante_impuestos_funcion[78,18] = comprobante_1[fila_comprobante_me_11, 11]
      
      #PAGO INTERESES DAMAS - Consecutivo 89:93
      #fila_comprobante_me_12 = 89
      comprobante_impuestos_funcion[99,19] = comprobante_1[fila_comprobante_me_12, 11]
      comprobante_impuestos_funcion[100,18] = comprobante_1[fila_comprobante_me_12, 11]
      
      #PAGO INTERESES FIJO DIARIO - Consecutivo 97:101
      #fila_comprobante_me_13 = 97
      comprobante_impuestos_funcion[101,19] = comprobante_1[fila_comprobante_me_13, 11]
      comprobante_impuestos_funcion[102,18] = comprobante_1[fila_comprobante_me_13, 11]
      
      #PAGO INTERESES CTA CTE - Consecutivo 105:109
      #fila_comprobante_me_14 = 105
      comprobante_impuestos_funcion[103,19] = comprobante_1[fila_comprobante_me_14, 11]
      comprobante_impuestos_funcion[104,18] = comprobante_1[fila_comprobante_me_14, 11]
      
      # Generar un nombre de variable con un índice
      index <- i
      var_name_1 <- paste("comprobante_impuestos_funcion_", index, sep = "")
      var_name_2 <- paste("Nombre_comprobante_impuestos_funcion_", index, sep = "")
      
      # Asignar la data al entorno global con el nombre dinámico
      assign(var_name_1, comprobante_impuestos_funcion, envir = .GlobalEnv)
      assign(var_name_2, var_name_1, envir = .GlobalEnv)
      
    },
    error = function(e){ 
      mensaje = paste('El archivo', comprobante_1, 'no esta disponible')
      cat(mensaje, "\n")
    }
  )
}


fun_comprobante_impuestos_funcion(comprobante_impuestos, 1,Hasta_1,Hasta_Mes_1, numero_semana_1,
                          1,9,17,25,73,81,33,49,57,113,121,129,89,97,105)

fun_comprobante_impuestos_funcion(comprobante_impuestos,2,Hasta_2,Hasta_Mes_2, numero_semana_2,
                          2,10,18,26,74,82,34,50,58,114,122,130,90,98,106)

fun_comprobante_impuestos_funcion(comprobante_impuestos,3,Hasta_3,Hasta_Mes_3, numero_semana_3,
                          3,11,19,27,75,83,35,51,59,115,123,131,91,99,107)

fun_comprobante_impuestos_funcion(comprobante_impuestos,4,Hasta_4,Hasta_Mes_4, numero_semana_4,
                          4,12,20,28,76,84,36,52,60,116,124,133,92,100,108)

fun_comprobante_impuestos_funcion(comprobante_impuestos,5,Hasta_5,Hasta_Mes_5, numero_semana_5,
                          5,13,21,29,77,85,37,53,61,117,125,134,93,101,109)

#--- exportación impuestos
#Comprobante_impuestos
wb = createWorkbook()
addWorksheet(wb, "Comprobante_impuestos_1")
addWorksheet(wb, "Comprobante_impuestos_2")
addWorksheet(wb, "Comprobante_impuestos_3")
addWorksheet(wb, "Comprobante_impuestos_4")
addWorksheet(wb, "Comprobante_impuestos_5")

tryCatch({writeData(wb, sheet = "Comprobante_impuestos_1", x = comprobante_impuestos_funcion_1, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "1", 'no esta disponible')
  cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Comprobante_impuestos_2", x = comprobante_impuestos_funcion_2, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "2", 'no esta disponible')
  cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Comprobante_impuestos_3", x = comprobante_impuestos_funcion_3, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "3", 'no esta disponible')
  cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Comprobante_impuestos_4", x = comprobante_impuestos_funcion_4, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "4", 'no esta disponible')
  cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Comprobante_impuestos_5", x = comprobante_impuestos_funcion_5, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "5", 'no esta disponible')
  cat(mensaje, "\n")})


saveWorkbook(wb, 'Comprobante_impuestos.xlsx', overwrite = TRUE)

drive_upload("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/Comprobante_impuestos.xlsx", path = as_id(folder_id_salida), name = "Comprobante_impuestos.xlsx", overwrite = TRUE)
#================================= Exportacion ============

#Creamos el libro
wb = createWorkbook()

#Creamos las hojas
addWorksheet(wb, "Comprobante")

#Comprobante Mary
writeData(wb, sheet = "Comprobante", x = "CONTRIBUCIÓN GRAVAMEN A LOS MOVIMIENTOS FINANCIEROS PAGADA POR EL BANCO", startCol = 1, startRow = 1)
writeData(wb, sheet = "Comprobante", x = fecha_completa, startCol = 1, startRow = 2)
writeData(wb, sheet = "Comprobante", x = comprobante_1, startCol = 1, startRow = 3)

#Formato comprobante
addStyle(wb,  sheet = "Comprobante", rows = 1, cols = 1:11, style = createStyle (textDecoration = "bold", fontColour = "#FF4040", fontSize = 14,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center",  borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 2, cols = 1:11, style = createStyle (textDecoration = "bold", fontColour = "#FF4040", fontSize = 13,  border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center",  borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 1, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 2, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 3, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 4, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 5, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 4:140, cols = 6, style = createStyle(textDecoration = "bold"))
addStyle(wb,  sheet = "Comprobante", rows = 1:140, cols = 11, style = createStyle( border = "right", borderColour = "black", borderStyle = "medium"))


addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 1, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 2, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 3, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 4, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 5, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113,121,129,137), cols = 6, style = createStyle (textDecoration = "bold", fgFill = "#FFFACD"))  #amarillo


addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 1, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 2, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 3, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 4, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 5, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 6, style = createStyle (textDecoration = "bold", fgFill = "#ADD8E6"))  #azul

addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 1, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 2, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 3, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 4, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 5, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 6, style = createStyle (textDecoration = "bold", fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa

addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 6, style = createStyle (fgFill = "#FFFACD", borderStyle = "medium"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 7, style = createStyle (fgFill = "#FFFACD", borderStyle = "medium"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 8, style = createStyle (fgFill = "#FFFACD", borderStyle = "medium"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 9, style = createStyle (fgFill = "#FFFACD", borderStyle = "medium"))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 10, style = createStyle (fgFill = "#FFFACD",border = "right", borderColour = "black", borderStyle = "medium" ))  #amarillo
addStyle(wb,  sheet = "Comprobante", rows = c(9,17,25,33,41,49,57,65,73,81,89,97,105,113, 121,129,137), cols = 11, style = createStyle (fgFill = "#FFFACD",border = "right", borderColour = "black", borderStyle = "medium" ))  #amarillo


addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 6, style = createStyle (fgFill = "#ADD8E6", borderStyle = "medium"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 7, style = createStyle (fgFill = "#ADD8E6",borderStyle = "medium"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 8, style = createStyle (fgFill = "#ADD8E6",borderStyle = "medium"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 9, style = createStyle (fgFill = "#ADD8E6",borderStyle = "medium"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 10, style = createStyle (fgFill = "#ADD8E6", border = "right", borderColour = "black", borderStyle = "medium"))  #azul
addStyle(wb,  sheet = "Comprobante", rows = c(10,18,26,34,42,50,58,66,74,82,90,98,106,114, 122,130,138), cols = 11, style = createStyle (fgFill = "#ADD8E6", border = "right", borderColour = "black", borderStyle = "medium"))  #azul

addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 6, style = createStyle (fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 7, style = createStyle (fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 8, style = createStyle (fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 9, style = createStyle (fgFill = "#FFE4E1", border = "bottom", borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 10, style = createStyle (fgFill = "#FFE4E1", border = c("bottom","right"), borderColour = "black", borderStyle = "medium"))  #rosa
addStyle(wb,  sheet = "Comprobante", rows = c(11,19,27,35,43,51,59,67,75,83,91,99,107,115, 123,131,139), cols = 11, style = createStyle (fgFill = "#FFE4E1", border = c("bottom","right"), borderColour = "black", borderStyle = "medium"))  #rosa


addStyle(wb,  sheet = "Comprobante", rows = 140, cols = 1:11, style = createStyle (textDecoration = "bold",  border = "bottom", borderColour = "black", borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 140, cols = 11, style = createStyle (textDecoration = "bold",  border = c("bottom", "right"), borderColour = "black", borderStyle = "medium"))

addStyle(wb,  sheet = "Comprobante", rows = 141, cols = 1:11, style = createStyle (textDecoration = "bold",  border = "bottom", borderColour = "black", borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 141, cols = 11, style = createStyle (textDecoration = "bold",  border = c("bottom", "right"), borderColour = "black", borderStyle = "medium"))

addStyle(wb,  sheet = "Comprobante", rows = 142, cols = 1:11, style = createStyle (textDecoration = "bold",  border = "bottom", borderColour = "black", borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 142, cols = 11, style = createStyle (textDecoration = "bold",  border = c("bottom", "right"), borderColour = "black", borderStyle = "medium"))

addStyle(wb,  sheet = "Comprobante", rows = 143, cols = 1:11, style = createStyle (textDecoration = "bold",  border = "bottom", borderColour = "black", borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 143, cols = 11, style = createStyle (textDecoration = "bold",  border = c("bottom", "right"), borderColour = "black", borderStyle = "medium"))

addStyle(wb,  sheet = "Comprobante", rows = 144, cols = 1:11, style = createStyle (textDecoration = "bold",  border = "bottom", borderColour = "black", borderStyle = "medium"))
addStyle(wb,  sheet = "Comprobante", rows = 144, cols = 11, style = createStyle (textDecoration = "bold",  border = c("bottom", "right"), borderColour = "black", borderStyle = "medium"))

addStyle(wb,  sheet = "Comprobante", rows = 3, cols = 1:11, style = createStyle(textDecoration = "bold", fontColour = "#0000FF", fontSize = 12, border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center", borderStyle = "medium"))

mergeCells(wb,  sheet = "Comprobante", cols = 1:11, rows = 1)   #Combinar celdas
mergeCells(wb,  sheet = "Comprobante", cols = 1:11, rows = 2)   #Combinar celdas

#Guardamos el libro
saveWorkbook(wb, 'Salidas.xlsx', overwrite = TRUE)

#Enviamos a la carpeta drive 
drive_upload("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/Salidas.xlsx", path = as_id(folder_id_salida), name = "Salidas.xlsx", overwrite = TRUE)


#CONTROLES
#Cheques girados
wb = createWorkbook()
addWorksheet(wb, "Control_CH_1")
addWorksheet(wb, "Control_CH_2")
addWorksheet(wb, "Control_CH_3")
addWorksheet(wb, "Control_CH_4")
addWorksheet(wb, "Control_CH_5")
addWorksheet(wb, "Control_CH_Completo")

tryCatch({writeData(wb, sheet = "Control_CH_1", x = R_cheques_girados_1, startCol = 1, startRow = 1) 
  writeData(wb, sheet = "Control_CH_1", x = cheques_girados_1, startCol = 1, startRow = 7)},error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_CH_2", x = R_cheques_girados_2, startCol = 1, startRow = 1) 
   writeData(wb, sheet = "Control_CH_2", x = cheques_girados_2, startCol = 1, startRow = 7)},error = function(e){ 
      mensaje = paste('El archivo', "2", 'no esta disponible')
      cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_CH_3", x = R_cheques_girados_3, startCol = 1, startRow = 1) 
  writeData(wb, sheet = "Control_CH_3", x = cheques_girados_3, startCol = 1, startRow = 7)},error = function(e){ 
    mensaje = paste('El archivo', "3", 'no esta disponible')
    cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_CH_4", x = R_cheques_girados_4, startCol = 1, startRow = 1) 
  writeData(wb, sheet = "Control_CH_4", x = cheques_girados_4, startCol = 1, startRow = 7)},error = function(e){ 
    mensaje = paste('El archivo', "4", 'no esta disponible')
    cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_CH_5", x = R_cheques_girados_5, startCol = 1, startRow = 1) 
  writeData(wb, sheet = "Control_CH_5", x = cheques_girados_5, startCol = 1, startRow = 7)},error = function(e){ 
    mensaje = paste('El archivo', "5", 'no esta disponible')
    cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_CH_Completo", x = R_cheques_girados_6, startCol = 1, startRow = 1) 
  writeData(wb, sheet = "Control_CH_Completo", x = cheques_girados_6, startCol = 1, startRow = 7)},error = function(e){ 
    mensaje = paste('El archivo', "6", 'no esta disponible')
    cat(mensaje, "\n")})
#Formato comprobante
addStyle(wb,  sheet = "Control_CH_1", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_1", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_1", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_1", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_1", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_1", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_CH_2", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_2", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_2", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_2", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_2", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_2", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_CH_3", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_3", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_3", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_3", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_3", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_3", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_CH_4", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_4", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_4", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_4", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_4", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_4", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_CH_5", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_5", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_5", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_5", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_5", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_5", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_CH_Completo", rows = 1, cols = 1:3, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_CH_Completo", rows = 2:5, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_Completo", rows = 2:5, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_Completo", rows = 2:5, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_Completo", rows = 5, cols = 1:3, style = createStyle (textDecoration = "bold",fgFill = "#FFD39B",  border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_CH_Completo", rows = 7, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

saveWorkbook(wb, 'Control_Cheques_Girados.xlsx', overwrite = TRUE)

drive_upload("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/Control_Cheques_Girados.xlsx", path = as_id(folder_id_salida), name = "Control_Cheques_Girados.xlsx", overwrite = TRUE)


#Cheques 
wb = createWorkbook()
addWorksheet(wb, "Control_Cheques_1")
addWorksheet(wb, "Enviar_a_Oficinas_1")
addWorksheet(wb, "Control_Cheques_2")
addWorksheet(wb, "Enviar_a_Oficinas_2")
addWorksheet(wb, "Control_Cheques_3")
addWorksheet(wb, "Enviar_a_Oficinas_3")
addWorksheet(wb, "Control_Cheques_4")
addWorksheet(wb, "Enviar_a_Oficinas_4")
addWorksheet(wb, "Control_Cheques_5")
addWorksheet(wb, "Enviar_a_Oficinas_5")
addWorksheet(wb, "Control_Cheques_Completo")
addWorksheet(wb, "Enviar_a_Oficinas_Completo")


tryCatch({writeData(wb, sheet = "Control_Cheques_1", x = Cheques_1, startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_1", x = Cheques_Oficina_1, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "1", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_Cheques_2", x = Cheques_2, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "2", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_2", x = Cheques_Oficina_2, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "2", 'no esta disponible')
  cat(mensaje, "\n")})


tryCatch({writeData(wb, sheet = "Control_Cheques_3", x = Cheques_3, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "3", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_3", x = Cheques_Oficina_3, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "3", 'no esta disponible')
  cat(mensaje, "\n")})


tryCatch({writeData(wb, sheet = "Control_Cheques_4", x = Cheques_4, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "4", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_4", x = Cheques_Oficina_4, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "4", 'no esta disponible')
  cat(mensaje, "\n")})


tryCatch({writeData(wb, sheet = "Control_Cheques_5", x = Cheques_5, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "5", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_5", x = Cheques_Oficina_5, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "5", 'no esta disponible')
  cat(mensaje, "\n")})

tryCatch({writeData(wb, sheet = "Control_Cheques_Completo", x = Cheques_6, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "6", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Enviar_a_Oficinas_Completo", x = Cheques_Oficina_6, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "6", 'no esta disponible')
  cat(mensaje, "\n")})


addStyle(wb,  sheet = "Control_Cheques_1", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_1", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_Cheques_2", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_2", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_Cheques_3", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_3", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_Cheques_4", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_4", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_Cheques_5", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_5", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))

addStyle(wb,  sheet = "Control_Cheques_Completo", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Enviar_a_Oficinas_Completo", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))


saveWorkbook(wb, 'Control_Cheques.xlsx', overwrite = TRUE)
drive_upload("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/Control_Cheques.xlsx", path = as_id(folder_id_salida), name = "Control_Cheques.xlsx", overwrite = TRUE)
#Pagos de intereses
wb = createWorkbook()
addWorksheet(wb, "Control_PIntereses_1")
addWorksheet(wb, "Control_PIntereses_2")
addWorksheet(wb, "Control_PIntereses_3")
addWorksheet(wb, "Control_PIntereses_4")
addWorksheet(wb, "Control_PIntereses_5")
addWorksheet(wb, "Control_PIntereses_Completo")

tryCatch({writeData(wb, sheet = "Control_PIntereses_1", x = Pago_Intereses_1, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "1", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_PIntereses_2", x = Pago_Intereses_2, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "2", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_PIntereses_3", x = Pago_Intereses_3, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "3", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_PIntereses_4", x = Pago_Intereses_4, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "4", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_PIntereses_5", x = Pago_Intereses_5, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "5", 'no esta disponible')
  cat(mensaje, "\n")})
tryCatch({writeData(wb, sheet = "Control_PIntereses_Completo", x = Pago_Intereses_6, startCol = 1, startRow = 1) },error = function(e){ 
  mensaje = paste('El archivo', "6", 'no esta disponible')
  cat(mensaje, "\n")})

addStyle(wb,  sheet = "Control_PIntereses_1", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_1", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_1", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_1", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_1", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_1", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))

addStyle(wb,  sheet = "Control_PIntereses_2", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_2", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_2", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_2", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_2", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_2", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))

addStyle(wb,  sheet = "Control_PIntereses_3", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_3", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_3", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_3", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_3", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_3", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))

addStyle(wb,  sheet = "Control_PIntereses_4", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_4", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_4", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_4", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_4", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_4", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))

addStyle(wb,  sheet = "Control_PIntereses_5", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_5", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_5", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_5", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_5", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_5", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))

addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 1, cols = 1:5, style = createStyle (textDecoration = "bold", fgFill = "#FF4040", fontSize = 12,   border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 2:4, cols = 1, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), borderColour = "black", halign = "center"))
addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 2:4, cols = 2, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 2:4, cols = 3, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 2:4, cols = 4, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
addStyle(wb,  sheet = "Control_PIntereses_Completo", rows = 2:4, cols = 5, style = createStyle (border = c("top", "bottom", "left", "right"), borderColour = "black"))
saveWorkbook(wb, 'Control_PIntereses.xlsx', overwrite = TRUE)
drive_upload("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/Control_PIntereses.xlsx", path = as_id(folder_id_salida), name = "Control_PIntereses.xlsx", overwrite = TRUE)

Funcion_Control_CDT = function(indice){ 
  encabezado_1 = "Formato Semanal a reportar"
  encabezado_2 = "Contribución"
  encabezado_3 = "INTERESES CERTIFICADOS  PESOS RENTABILIDAD CDTS"
  
  wb = createWorkbook()
  addWorksheet(wb, "Resumen")
  addWorksheet(wb, "DECEVAL_DAVIVIENDA")
  addWorksheet(wb, "Anulacion_VS_Apertura")
  addWorksheet(wb, "Intereses")
  addWorksheet(wb, "TD_14919_vs_14903")
  addWorksheet(wb, "14919_VS_14903")
  addWorksheet(wb, "TD_14901vs14919")
  addWorksheet(wb, "14901_VS_14919")
  addWorksheet(wb, "Resumen_14903")
  addWorksheet(wb, "TD_contribucion")
  addWorksheet(wb, "Fondo_de_inversion")
  addWorksheet(wb, "Rtran_VS_redenciones")

  
  writeData(wb, sheet = "Resumen", x = encabezado_1, startCol = 1, startRow = 1)
  writeData(wb, sheet = "Resumen", x = encabezado_2, startCol = 1, startRow = 12)
  writeData(wb, sheet = "Resumen", x = encabezado_3, startCol = 1, startRow = 21)
  tryCatch({writeData(wb, sheet = "Resumen", x = get(paste0("Formato_semanal_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 2) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Resumen", x = get(paste0("Contribucion_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 13) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Resumen", x = get(paste0("Intereses_Certificados_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 22) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "DECEVAL_DAVIVIENDA", x = get(paste0("DECEVAL_DAVIVIENDA_",indice), envir = .GlobalEnv), startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "2", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Anulacion_VS_Apertura", x = get(paste0("Anulacion_VS_Apertura_",indice), envir = .GlobalEnv), startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "3", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Intereses", x = get(paste0("Intereses_",indice), envir = .GlobalEnv), startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "4", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "TD_14919_vs_14903", x = get(paste0("tabla_dinamica_14919_vs_14903_",indice), envir = .GlobalEnv), startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "4", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "14919_VS_14903", x = get(paste0("14919_VS_14903_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "TD_14901vs14919", x = get(paste0("tabla_dinamica_14901vs14919_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "14901_VS_14919", x = get(paste0("14901_VS_14919_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Resumen_14903", x = get(paste0("Resumen_14903_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "TD_contribucion", x = get(paste0("tabla_dinamica_contribucion_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "1", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Fondo_de_inversion", x = get(paste0("Cruza_fondos_14903_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "Cruce_fondo_inversion_", 'no esta disponible')
    cat(mensaje, "\n")})
  tryCatch({writeData(wb, sheet = "Rtran_VS_redenciones", x = get(paste0("Cruza_rtran_redenciones_",indice), envir = .GlobalEnv) , startCol = 1, startRow = 1) },error = function(e){ 
    mensaje = paste('El archivo', "Cruza_rtran_redenciones_", 'no esta disponible')
    cat(mensaje, "\n")})
le estilo a las salidas
  addStyle(wb,  sheet = "Resumen", rows = 1, cols = 1, style = createStyle (textDecoration = "bold", fontSize = 12,  halign = "center"))
  addStyle(wb,  sheet = "Resumen", rows = 12, cols = 1, style = createStyle (textDecoration = "bold", fontSize = 12, halign = "center"))
  addStyle(wb,  sheet = "Resumen", rows = 21, cols = 1, style = createStyle (textDecoration = "bold", fontSize = 12, halign = "center"))
  mergeCells(wb,  sheet = "Resumen", cols = 1:5, rows = 1)  
  mergeCells(wb,  sheet = "Resumen", cols = 1:6, rows = 12)  
  mergeCells(wb,  sheet = "Resumen", cols = 1:5, rows = 21)  
  
  addStyle(wb,  sheet = "Resumen", rows = 2, cols = 1:5, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Resumen", rows = 13, cols = 1:6, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white", fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Resumen", rows = 22, cols = 1:5, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white", fgFill = "#FF4040", borderColour = "black", halign = "center"))
  
  addStyle(wb,  sheet = "DECEVAL_DAVIVIENDA", rows = 1, cols = 1:51, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Anulacion_VS_Apertura", rows = 1, cols = 1:51, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Intereses", rows = 1, cols = 1:51, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "TD_14919_vs_14903", rows = 1, cols = 1:4, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "14919_VS_14903", rows = 1, cols = 1:51, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "TD_14901vs14919", rows = 1, cols = 1:4, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Resumen_14903", rows = 1, cols = 1:4, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "TD_contribucion", rows = 1, cols = 1:6, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Fondo_de_inversion", rows = 1, cols = 1:55, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  addStyle(wb,  sheet = "Rtran_VS_redenciones", rows = 1, cols = 1:69, style = createStyle (textDecoration = "bold", border = c("top", "bottom", "left", "right"), fontColour = "white" , fgFill = "#FF4040", borderColour = "black", halign = "center"))
  
  
  saveWorkbook(wb, paste0("Control_CDT_Semana_", indice, ".xlsx"), overwrite = TRUE)
  
  directorio = paste0("D:/Mis Documentos/GMF/Comprobante_Completo/ENTRADA/","Control_CDT_Semana_", indice,".xlsx")
  nombre = paste0("Control_CDT_Semana_", indice, ".xlsx")
  drive_upload(directorio, path = as_id(folder_id_salida), name = nombre, overwrite = TRUE)
  
}

Funcion_Control_CDT(1)
Funcion_Control_CDT(2)
Funcion_Control_CDT(3)
Funcion_Control_CDT(4)
Funcion_Control_CDT(5)
Funcion_Control_CDT(6)
