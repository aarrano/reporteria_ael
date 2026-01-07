# ============================================================================
# SCRIPT PARA ENVÍO AUTOMATIZADO DE CORREOS CON OUTLOOK
# Usando PowerShell desde R (sin necesidad de contraseñas)
# ============================================================================

# 1. LIBRERIAS ----
rm(list=ls())
pacman::p_load(tidyverse,openxlsx,readxl)


# 2. POWERSHELL ----
if (Sys.info()["sysname"] != "Windows") {
  stop("Este script requiere Windows con PowerShell")
}

cat("Verificando PowerShell...\n")
test_ps <- system2("powershell", args = c("-Command", "Write-Output 'OK'"), 
                   stdout = TRUE, stderr = TRUE)
if (test_ps[1] != "OK") {
  stop("PowerShell no está disponible o no funciona correctamente")
}
cat("✓ PowerShell disponible\n\n")

# 3. LEER DATOS DEL EXCEL ----

# Ruta del archivo Excel
ruta_excel <- "./correos/correos_prueba.xlsx"

# Carpeta donde están los PDFs
carpeta_pdfs <- "./Minuta x SLEP/2026/260105/"

# Leer datos
df <- read.xlsx(ruta_excel) %>% rename_all(tolower)

link_archivo <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/reporteria_ael/Minuta x SLEP/2026/"
fecha_archivo <- "260105/"
path_archivo <- paste0(link_archivo,fecha_archivo)

df <- df %>%
  mutate(
    correo = ifelse(is.na(correo), "", trimws(correo)),
    copia  = ifelse(is.na(copia), "", trimws(copia)),
    slep   = ifelse(is.na(slep), "Sin Nombre", trimws(slep)),
    nom_archivo = trimws(archivo) # Elimina espacios al inicio o final
  ) %>% 
  mutate(archivo=paste0(path_archivo,nom_archivo)) %>% 
  mutate(saludo = case_when(
    genero == "F" ~ "Estimada ",
    genero == "M" ~ "Estimado "
  ))

# Correos nuevos ----
df <- df %>% 
  filter(slep == "Petorca")


# Estructura esperada:
# - empresa: Nombre de la empresa
# - nombre: Nombre de la persona
# - correo: Email del destinatario
# - archivo_pdf: Nombre del archivo PDF

cat("Vista previa de los datos:\n")
print(head(df))
cat("\nTotal de registros:", nrow(df), "\n\n")

# Función ENVIO DE CORREOS ----
enviar_outlook_ps <- function(destinatario, cc, asunto, cuerpo, ruta_archivo) {
  
  ruta_limpia <- normalizePath(ruta_archivo, winslash = "\\", mustWork = FALSE)
  
  if (!file.exists(ruta_limpia)) {
    warning(paste("Archivo no encontrado:", ruta_limpia))
    return(NULL)
  }
  
  # Construir la línea de CC solo si hay un correo
  cc_line <- if(cc != "") sprintf('$Mail.CC = "%s";', cc) else ""
  
  # Comando PowerShell
  ps_cmd <- sprintf(
    '$Outlook = New-Object -ComObject Outlook.Application; 
     $Mail = $Outlook.CreateItem(0); 
     $Mail.To = "%s"; 
     %s 
     $Mail.Subject = "%s"; 
     $Mail.Body = "%s"; 
     $Mail.Attachments.Add(\'%s\') | Out-Null; 
     $Mail.Send()', 
    destinatario, cc_line, asunto, cuerpo, ruta_limpia
  )
  
  # Ejecutar
  resultado <- system2("powershell", 
                       args = c("-NoProfile", "-ExecutionPolicy Bypass", "-Command", 
                                shQuote(enc2utf8(ps_cmd))),
                       stdout = TRUE, stderr = TRUE)
  
  # Si hay error en PowerShell, mostrarlo en R
  if (length(resultado) > 0 && any(grepl("Error", resultado))) {
    message("Error detectado en el envío:")
    print(resultado)
  }
}

# LOOP ENVIO ----
for (i in 1:nrow(df)) {
  
  nombre_cli  <- df$nombre[i]
  email_cli   <- df$correo[i]
  copia_cli <- df$copia[i]
  empresa_cli <- df$slep[i]
  archivo     <- df$archivo[i]
  saludo_cli <- df$saludo[i]
  
  asunto_msg <- paste("Reporte Semanal AEL -", empresa_cli)
  cuerpo_msg <- paste0(
    saludo_cli, nombre_cli, ",\n\n",
    "Junto con saludar y desearles un feliz año, les escribo para compartir la actualización del estado del 'Anótate en la lista' al 05 de enero para el SLEP ", empresa_cli, ".\n\n",
    "En el archivo podrán encontrar los datos sobre listas de espera, establecimientos sin cuentas activas y otros, por lo que les sugerimos se centren en aquellos establecimientos sin cuentas activas y con lista de espera.\n\n",
    "Adicionalmente, priorizar el contacto con aquellos establecimientos con una alta cantidad de Vacantes Sin Asignar (dado que es una forma de aumentar la matrícula en algunos EE).\n\n",
    "Veremos este instrumento en mayor detalle en la reunión del Martes 13 de Enero, donde tendremos la inducción sobre el Sistema de Admisión Escolar y las labores asociadas.\n\n",
    "Saludos cordiales.\n\n",
    "PD: Este es un correo generado automáticamente, favor reportar cualquier incidencia con el envío de información.\n\n",
    "PD 2: El archivo debe ser descargado y visualizado desde un navegador, no es compatible con la vista desde el telefóno .\n\n"
  )
  # Validacion que existe el archivo
  if (!file.exists(archivo)) {
    message(paste("EERR: Archivo no encontrado para", empresa_cli, "- Saltando..."))
    next # Salta a la siguiente fila del Excel
  }
  
  # Llamar a la función
  enviar_outlook_ps(email_cli,copia_cli,asunto_msg, cuerpo_msg, archivo)
  
  message(paste("Correo procesado para:", empresa_cli))
  Sys.sleep(2) # Pausa breve para no saturar Outlook
}
