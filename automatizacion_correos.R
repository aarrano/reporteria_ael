# ============================================================================
# SCRIPT PARA ENVÍO AUTOMATIZADO DE CORREOS CON OUTLOOK
# Usando PowerShell desde R 


# 1. LIBRERIAS ----
rm(list=ls())
pacman::p_load(tidyverse,openxlsx,readxl)

# 2. CARGAR TEMPLATES DE CORREOS ----
# Los mensajes HTML se mantienen en un archivo separado para facilitar edición
source("template_correos.R")

# 3. POWERSHELL ----
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

# 4. LEER DATOS DEL EXCEL ----

fecha_archivo <- "260209/"

# Ruta del archivo Excel
ruta_excel <- "./correos/correos_FINAL.xlsx"

# Carpeta donde están los PDFs
carpeta_pdfs <- paste0("./Minuta x SLEP/2026/",fecha_archivo)

# Leer datos
df <- read.xlsx(ruta_excel) %>% rename_all(tolower)

link_archivo <- "D:/Alonso.Arrano/OneDrive - Dirección de Educación Pública/2024/SAE - Anotate en la lista/reporteria_ael/Minuta x SLEP/2026/"

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

# Estructura esperada:
# - empresa: Nombre de la empresa
# - nombre: Nombre de la persona
# - correo: Email del destinatario
# - archivo_pdf: Nombre del archivo PDF

cat("Vista previa de los datos:\n")
print(head(df))
cat("\nTotal de registros:", nrow(df), "\n\n")

# Función ENVIO DE CORREOS ----
enviar_outlook_ps <- function(destinatario, cc, asunto, cuerpo_html, ruta_archivo) {
  
  ruta_limpia <- normalizePath(ruta_archivo, winslash = "\\", mustWork = FALSE)
  
  if (!file.exists(ruta_limpia)) {
    warning(paste("Archivo no encontrado:", ruta_limpia))
    return(NULL)
  }
  
  # Escapar comillas dobles en el HTML para PowerShell
  cuerpo_escapado <- gsub('"', '`"', cuerpo_html)
  
  # Construir la línea de CC solo si hay un correo
  cc_line <- if(cc != "") sprintf('$Mail.CC = "%s";', cc) else ""
  
  # Comando PowerShell
  ps_cmd <- sprintf(
    '$Outlook = New-Object -ComObject Outlook.Application; 
     $Mail = $Outlook.CreateItem(0); 
     $Mail.To = "%s"; 
     %s 
     $Mail.Subject = "%s"; 
    $Mail.HTMLBody = "%s"; 
     $Mail.Attachments.Add(\'%s\') | Out-Null; 
     $Mail.Send()', 
    destinatario, cc_line, asunto, cuerpo_escapado, ruta_limpia
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

z=1

# Testear el cuerpo del correo:
# writeLines(cuerpo_html, "preview.html")
# shell.exec("preview.html")

# LOOP ENVIO ----
for (i in 1:nrow(df)) {
  
  nombre_cli  <- df$nombre[i]
  email_cli   <- df$correo[i]
  copia_cli <- df$copia[i]
  empresa_cli <- df$slep[i]
  archivo     <- df$archivo[i]
  saludo_cli <- df$saludo[i]
  
  asunto_msg <- paste("Reporte Semanal AEL -", empresa_cli)

  # El cuerpo se genera usando el codigo de template_correo.R
  cuerpo_html <- generar_html_reporte_semanal(
    saludo = saludo_cli,
    nombre = nombre_cli,
    empresa = empresa_cli
  )
  
  # Validacion que existe el archivo
  if (!file.exists(archivo)) {
    message(paste("ERROR: Archivo no encontrado para", empresa_cli, "- Saltando..."))
    next # Salta a la siguiente fila del Excel
  }
  
  # Llamar a la función
  enviar_outlook_ps(email_cli,copia_cli,asunto_msg, cuerpo_html, archivo)
  
  message(paste("Correo procesado para:", empresa_cli, "(",z," de ",nrow(df),")"))
  z=z+1
  Sys.sleep(2) # Pausa breve para no saturar Outlook
}
