#Autor: Alonso Arraño Portuguez
#Fecha: 06-10-2025
#Objetivo: Generar reportería automatizada respecto al proceso de anótate en la lista.
#A modo de complementar el traspaso de información de la DEP hacia los SLEP.
#Solicitante: N/A
#Motivo: Este proyecto surge del interés personal de apoyar la gestión de los SLEP
# a través de la provisión de información.
#Sobretodo pensado que el visor de powerbi tuvo su auge de uso durante los primeros 
# 4 meses del año. Así, logramos instalar un proceso clave en la UATP. 

#Contexto: Originalmente habia un seguimiento semanal que se hacia con una minuta y datos manuales.
#Por falta de prioridad el proceso nunca se llevó a cabo, hasta ahora, que hay más tiempo disponible.
rm(list=ls())

# Proyecto ----
message("Inicio proceso creación de reportes AEL para SLEP")

pacman::p_load(tidyverse,data.table,openxlsx,janitor,gt,fontawesome,knitr,quarto,ggrepel)

year_actual <- year(today())

## Pathways ---

#Acceso al maestro del AEL
link_maestro <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/output/df_maestro_ael_2025.csv"

#Acceso al ultimo AEL
link_1 <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/output/datos visualizador 24_10_28/"
files <- list.files(link_1,full.names = TRUE)

ael_t <- fread(files[length(files)-1])

message(paste("Leyendo el archivo",files[length(files)-1],"como archivo actual"))

#Acceso al estado de cuentas 
link_cuentas <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/output/Cuentas Activas/cuentas_activas.xlsx"

#Cargamos tabla con definiciones para el glosario
tabla_glosario <- read.xlsx("./inputs_qmd/tabla_glosario.xlsx")


# Carga de información ----

df_ael <- fread(link_maestro)
df_cuentas <- read.xlsx(link_cuentas) %>%  clean_names()
df_cuentas <- df_cuentas%>%
  filter(d_inactivo==0) 

## Trabajando información ----

# Añadimos un recordatorio para que avise el código si nos equivocamos en la cantidad de SLEP
if (year_actual != "2025") {
  stop("❌ Debes actualizar el código para que la generación de reportes incluya a los SLEP que están en su primer año de instalación.")
} else {
  message("✅ El código se encuentra actualizado para los SLEP en régimen.")
}

# Dejamos solo los 26 SLEP actuales
df_ael <- df_ael %>% filter(!nombre_slep %in% c("Marga Marga","Tamarugal")) 
nombre_sleps <- unique(df_ael$nombre_slep)
print(length(nombre_sleps))

## Definimos bases para filtrar ----
message("Debemos generar distintas bases para poder acceder a los distintos indicadores que generan el reporte, ideal para que el loop no tome tanto tiempo")

### 1 - Indicadores Tier 1 ----
"Esta sección incluye los 5 indicadores mas utilizados: Total de EE, EE atrasados, % de cumplimiento, Prom. dias sin movimiento, EE sin cuentas activas"

fecha_ultima_actualización <- max(df_ael$fecha_corte_info)
print(paste("Los últimos datos corresponden a",fecha_ultima_actualización))

indicadores_1 <- df_ael %>%
  filter(fecha_corte_info == fecha_ultima_actualización) %>%
  summarize(`Total de EE` = n_distinct(rbd),
            `EE atrasados` = n_distinct(rbd[condicion_rbd == 1]),
            `Tasa de cumplimiento` = round(`EE atrasados`*100/`Total de EE`,1),
            `Promedio dias atrasados` = round(mean(`promedio dias sin movimiento`[condicion_rbd == 1 & posibles_cupos>0]),1),
            `Vacantes sin asignar` = sum(posibles_cupos, na.rm = TRUE),
            .by = nombre_slep) %>% 
  left_join(df_cuentas %>%
              summarize(`EE sin cuentas`=n(),.by = nombre_slep),
            by = "nombre_slep") %>% 
  mutate(`Tasa EE sin cuenta` = round(`EE sin cuentas`*100/`Total de EE`,1)) %>% 
  mutate(`Tasa EE sin cuenta` = if_else(is.na(`EE sin cuentas`) ,0,`Tasa EE sin cuenta`))

# Reemplazamos con 0 los SLEP que no tienen cuentas atrasadas
indicadores_1[is.na(indicadores_1)] <- 0


### 2 - Listado de EE sin cuenta activa ----
"Generamos la base de cuentas inactivas por SLEP, además añadiremos la cantidad de lista de espera de cada RBD"

df_cuentas <- df_cuentas %>% 
  left_join(df_ael %>%
              filter(fecha_corte_info == fecha_ultima_actualización) %>%
              summarize(`Lista de espera` = sum(lista_de_espera_anotate,na.rm=TRUE),
                        `Vacantes disponibles` = sum(vacantes_para_analisis, na.rm = TRUE),
                        .by = rbd) %>% 
              mutate(rbd = as.character(rbd))
              , by = "rbd")

"Algunas variables quedan con NA porque no hay correos asociados al RBD en SIGE."
df_cuentas[is.na(df_cuentas)] <- ""

#### 2.1 Base de cuentas inactivas para el reporte ----

"Seleccionamos variables de interés de la base df_cuentas"
"Realizamos un cambio en el formato de la fecha, que al importar se lee como string"
"Pasamos la fecha a formato dia- mes - año"

df_cuentas_slep <- df_cuentas %>% 
  select(nombre_slep,
         #comuna,
         establecimiento,`Lista de espera`,correo_director,correo_encargado_sae,starts_with("fecha_creacion_")) %>% 
  mutate( fecha_creacion_director = format(
    as.POSIXct(fecha_creacion_director, format = "%Y-%m-%d %H:%M:%S"),
    "%d-%m-%Y"
  ),
  fecha_creacion_encargado_sae = format(
    as.POSIXct(fecha_creacion_encargado_sae, format = "%Y-%m-%d %H:%M:%S"),
    "%d-%m-%Y"
  ))

### 3 - Grafico AEL histórico por comuna ----

"Generamos una base con el histórico de Vacantes sin Asignar, por fecha y comuna"
temp_1 <- df_ael %>%
  summarize(vas = sum(posibles_cupos,na.rm = TRUE),.by = c(nombre_slep,fecha_corte_info,comuna)) %>% 
  mutate(comuna2=comuna)

message("Con los distintos DF puntuales creados, se procede al loop para generar los reportes de cada SLEP")

### 4 - Listado AEL actual por establecimiento 

temp_ael <- ael_t %>% 
  filter(condicion_rbd==1) %>% 
  filter(id == 1) %>% 
  arrange(id_orden_rbd) %>% 
  select(nombre_slep,comuna,rbd,nombre_ee,`total posibles cupos`,`total niveles atrasados`,`total niveles ofrecidos`,`promedio dias sin movimiento`) %>% 
  rename(`Vacantes sin asignar` = `total posibles cupos`)

# LOOP por SLEP ----

for (s in nombre_sleps[25]) {
  
  "Hacemos el print de qué SLEP se está generando"
  print(paste("Trabajando en el slep", s))
  
  ## Pasamos los indicadores claves del SLEP ----
  n_ee <- indicadores_1[nombre_sleps == s, 2]
  n_ee_atrasados <-  indicadores_1[nombre_sleps == s, 3]
  n_ee_sin_cuenta <-  indicadores_1[nombre_sleps == s, 7]
  tasa_cumplimiento <-  indicadores_1[nombre_sleps == s, 4]
  dias_sin_mov <-  indicadores_1[nombre_sleps == s, 5]
  vas <- indicadores_1[nombre_sleps == s, 6]
  nombre = s
  
  ## Pasamos la lista de correos sin AEL ----
'Para evitar problemas con las tablas vacias se añadió el paso que deja los NA en "" una vez filtrada la tabla '
  
  df_cuentas_slep_qmd <- df_cuentas_slep %>% 
    filter(nombre_slep == s) %>% 
    select(-nombre_slep)
  
  df_cuentas_slep_qmd[is.na(df_cuentas_slep_qmd)] <- ""
  
  ## Pasamos el trend line por comuna del SLEP ----
  "Filtramos la base con los datos longitudinales por comuna"
  temp_2 <- temp_1 %>% 
    filter(nombre_slep == s)
  
  ## Pasamos el estado AEL de cada RBD ----
  ael_actual <- temp_ael %>% 
    filter(nombre_slep == s) %>% 
    select(-nombre_slep)

  ## Quarto render ----
  
  quarto::quarto_render(
    input = "./code/reporteria_ael_slep.qmd",
    execute_dir = getwd(), 
    output_format = "html",
    output_file = paste0("reporte_", gsub(" ", "_", s), ".html"),
    execute_params = list(
      slep = nombre,
      n_ee = n_ee,
      n_ee_atrasados = n_ee_atrasados,
      n_ee_sin_cuenta = n_ee_sin_cuenta,
      tasa_cumplimiento = tasa_cumplimiento,
      dias_sin_mov = dias_sin_mov,
      listado_cuentas = df_cuentas_slep_qmd,
      data_fig_1 = indicadores_1,
      graf_lineas = temp_2,
      vas = vas,
      glosario = tabla_glosario,
      ael_actual = ael_actual
    ),
    quiet = TRUE,
    quarto_args = c("--output-dir", "../Minuta x SLEP./251027/"),
  )

#Fin del loop
}




