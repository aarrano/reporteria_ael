# ============================================================================
# SCRIPT PARA ENVÍO AUTOMATIZADO DE CORREOS CON OUTLOOK
# Usando PowerShell desde R 


# 1. LIBRERIAS ----
rm(list=ls())
pacman::p_load(tidyverse,openxlsx,readxl)

# 1.1 CONFIGURACION ----
MODO_PRUEBA <- TRUE # CAMBIAR A FALSE CUANDO ESTÉ LISTO

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

fecha_archivo <- "260302/" # <--- DEBES MODIFICAR LA FECHA!!!!

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
    archivo1 = ifelse(is.na(archivo1), "", trimws(archivo1)),
    archivo2 = ifelse(is.na(archivo2), "", trimws(archivo2))
  ) %>% 
  mutate(ruta_archivo1 = ifelse(archivo1 != "", paste0(path_archivo, archivo1), ""),
         ruta_archivo2 = ifelse(archivo2 != "", paste0(path_archivo, archivo2), "")) %>% 
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
enviar_outlook_html <- function(destinatario, cc, asunto, cuerpo_html, archivos_adjuntos, modo_prueba = T) {
  
  # archivos_adjuntos puede ser:
  # - Una ruta única: "C:/archivo.pdf"
  # - Vector de rutas: c("C:/archivo.pdf", "C:/datos.xlsx")
  
  # Convertir a vector si es una sola ruta
  if (is.character(archivos_adjuntos) && length(archivos_adjuntos) == 1) {
    archivos_adjuntos <- c(archivos_adjuntos)
  }
  
  # Validar que todos los archivos existen
  archivos_limpio <- c()
  for (archivo in archivos_adjuntos) {
    ruta_limpia <- normalizePath(archivo, winslash = "\\", mustWork = FALSE)
    
    if (!file.exists(ruta_limpia)) {
      warning(paste("Archivo no encontrado:", ruta_limpia))
      return(NULL)
    }
    archivos_limpio <- c(archivos_limpio, ruta_limpia)
  }
  
  # Escapar comillas dobles en el HTML para PowerShell
  cuerpo_escapado <- gsub('"', '`"', cuerpo_html)
  
  # Construir la línea de CC solo si hay un correo
  cc_line <- if(cc != "") sprintf('$Mail.CC = "%s";', cc) else ""
  
  # Construir líneas para adjuntar múltiples archivos
  attachments_lines <- paste(
    sapply(archivos_limpio, function(f) {
      sprintf('$Mail.Attachments.Add(\'%s\') | Out-Null;', f)
    }),
    collapse = " "
  )
  
  # Comando PowerShell - Cambiar .Send() por .Display() en modo prueba
  if (modo_prueba) {
    # En modo prueba: abre el correo en Outlook para revisión (NO lo envía)
    enviar_comando <- '$Mail.Display()'
  } else {
    # En modo normal: envía automáticamente
    enviar_comando <- '$Mail.Send()'
  }
  
  ps_cmd <- sprintf(
    '$Outlook = New-Object -ComObject Outlook.Application; 
     $Mail = $Outlook.CreateItem(0); 
     $Mail.To = "%s"; 
     %s 
     $Mail.Subject = "%s"; 
     $Mail.HTMLBody = "%s"; 
     %s
     %s', 
    destinatario, cc_line, asunto, cuerpo_escapado, attachments_lines, enviar_comando
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
for (i in 1:nrow(df)[1]) {
  
  nombre_cli  <- df$nombre[i]
  email_cli   <- df$correo[i]
  copia_cli <- df$copia[i]
  empresa_cli <- df$slep[i]
  saludo_cli <- df$saludo[i]
  
  #Archivo PDF y excel
  archivo1 <- df$ruta_archivo1[i]
  archivo2 <- df$ruta_archivo2[i]
  
  # Crear vector de archivos (solo los que existen)
  archivos_adjuntos <- c()
  if (archivo1 != "") archivos_adjuntos <- c(archivos_adjuntos, archivo1)
  if (archivo2 != "") archivos_adjuntos <- c(archivos_adjuntos, archivo2)
  
  # Validar que hay al menos un archivo
  if (length(archivos_adjuntos) == 0) {
    message(paste("ERROR: No hay archivos para adjuntar a", empresa_cli, "- Saltando..."))
    next
  }
  
  # Validar que los archivos existen
  archivos_validos <- c()
  for (arch in archivos_adjuntos) {
    if (file.exists(arch)) {
      archivos_validos <- c(archivos_validos, arch)
    } else {
      warning(paste("Archivo no encontrado:", arch))
    }
  }
  
  if (length(archivos_validos) == 0) {
    message(paste("ERROR: Ningún archivo válido para", empresa_cli, "- Saltando..."))
    next
  }
  
  asunto_msg <- paste("Reporte Semanal AEL -", empresa_cli)

  # El cuerpo se genera usando el codigo de template_correo.R
  cuerpo_html <- generar_html_reporte_semanal(
    saludo = saludo_cli,
    nombre = nombre_cli,
    empresa = empresa_cli
  )
  
  
  # Enviar correo
  enviar_outlook_html(
    destinatario = email_cli, 
    cc = copia_cli, 
    asunto = asunto_msg, 
    cuerpo_html = cuerpo_html, 
    archivos_adjuntos = archivos_validos,
    modo_prueba = MODO_PRUEBA
  )
  
  if (MODO_PRUEBA) {
    message(paste("✉️  Correo creado (PRUEBA) para:", empresa_cli, "(", z, " de ", nrow(df), ")"))
  } else {
    message(paste("✅ Correo enviado para:", empresa_cli, "(", z, " de ", nrow(df), ")"))
  }
  
  z=z+1
  Sys.sleep(2) # Pausa breve para no saturar Outlook
}
