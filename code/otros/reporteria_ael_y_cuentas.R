#Autor: Alonso Arraño Portuguez
#Fecha: 25-03-25
#Objetivo: Automatizar el estado AEL + EE sin Cuentas activas

# Paquetes ----

pacman::p_load(tidyverse,data.table,openxlsx,janitor,writexl)

# Directorio ----

link_ael <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/output/datos visualizador 24_10_28/"
link_cuentas <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/output/Cuentas Activas/"

# Listamos las bbdd disponibles en la carpeta ael

# Obtener la lista de archivos y su información ----


files_info <- file.info(list.files(path =link_ael, full.names = TRUE))

# Ordenar los archivos por fecha de modificación
files <- rownames(files_info[order(files_info$mtime, decreasing = TRUE), ])
files




## Cargamos BBDD de estado AEL ----

bbdd_ael <- fread(files[2]) %>% rename_all(tolower) 

### Modificamos BBDD AEL ----

bbdd_ael <- bbdd_ael %>% 
  select(comuna,nombre_slep,rbd,nombre_ee,vacantes_para_analisis,lista_de_espera_anotate,posibles_cupos,fecha1) %>% 
  filter(!nombre_slep %in% c("Costa Central", "Marga Marga", "Tamarugal" , "Elqui") )

bbdd_ael_rbd <- bbdd_ael %>% 
  summarize(comuna=first(comuna),
            nombre_slep = first(nombre_slep),
            vacantes_para_analisis = sum(ifelse(vacantes_para_analisis>=0,vacantes_para_analisis,0)),
            total_lista_espera = sum(lista_de_espera_anotate),
            posibles_cupos = sum(posibles_cupos),
            ult_vacante_asignada = last(fecha1),
            .by = c(rbd,nombre_ee))
  
## Cargamos BBDD cuentas Activas AEL ----

bbdd_cuentas_activas <- read.xlsx(paste0(link_cuentas,"cuentas_activas.xlsx")) %>% rename_all(tolower) %>% clean_names()

### Modificamos BBDD cuentas_activas ----

bbdd_cuentas_activas <- bbdd_cuentas_activas %>% 
  select(establecimiento,ee_inactivo) %>% 
  mutate(rbd = as.integer(str_extract(establecimiento, "(?<=- )\\d+$")) ) %>% 
  filter(!is.na(rbd)) %>% 
  rename(EE_sin_cuenta_activa = ee_inactivo) %>% 
  select(-establecimiento)

## Pegamos información de ambos ----

bbdd_final <- bbdd_ael_rbd %>% 
  left_join(bbdd_cuentas_activas,by = "rbd") %>% 
  mutate(prioridad = case_when(EE_sin_cuenta_activa == "EE Inactivo" & posibles_cupos >0 ~ "(Urgente) no tiene cuenta y tiene lista de espera",
                               EE_sin_cuenta_activa == "EE Inactivo" & posibles_cupos <=0 ~ "Debe crear cuenta",
                               TRUE ~ "NA")) %>% 
  rename(vacantes_disponibles = vacantes_para_analisis) %>% 
  select(nombre_slep,everything())
         
nombre_slep <- unique(bbdd_final$nombre_slep)

fecha_formateada <- format(as.Date(today()), "%y_%m_%d")

for (nombre in nombre_slep) {
  df <- bbdd_final %>% filter((nombre_slep == nombre))
  
  write_xlsx(df,paste0("00. Reporteria/Envios SLEP/",fecha_formateada,"/estado_ael_cuentas_slep_",nombre,".xlsx"))
}
