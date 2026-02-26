# ============================================================================
# TEMPLATES HTML PARA CORREOS
# Archivo de configuración de mensajes
# ============================================================================

# Este archivo contiene todas las funciones para generar el cuerpo HTML
# de los diferentes tipos de correos que se envían.


# TEMPLATE: Reporte Semanal AEL ----
generar_html_reporte_semanal <- function(saludo, nombre, empresa) {
  
  html <- paste0('
  <html>
  <head>
    <meta charset="UTF-8">
    <style>
      body {
        font-family: Calibri, Arial, sans-serif;
        font-size: 11pt;
        color: #333;
        line-height: 1.6;
      }
      .contenedor {
        max-width: 700px;
        margin: 0 auto;
        padding: 20px;
      }
      .saludo {
        margin-bottom: 15px;
        font-size: 11pt;
      }
      .parrafo {
        margin: 12px 0;
        text-align: justify;
      }
      .destacado {
        text-decoration: underline;
        font-weight: bold;
      }
      .negrita {
        font-weight: bold;
      }
      .firma {
        margin-top: 25px;
        color: #555;
        font-size: 10pt;
      }
      .nota {
        margin-top: 20px;
        padding: 10px;
        background-color: #f0f0f0;
        border-left: 4px solid #0078D4;
        font-size: 10pt;
        font-style: italic;
      }
    </style>
  </head>
  <body>
    <div class="contenedor">
      <p class="saludo">', saludo, ' ', nombre, ',</p>
      
      <p class="parrafo">
        Junto con saludar, y esperando se encuentre bien, comparto con usted 
        <span class="destacado">la actualización del estado del \'Anótate en la lista\' 
        al 16 de febrero </span> para el SLEP ', empresa, '.
      </p>
      
      <p class="parrafo">
        En el archivo podrán encontrar los datos sobre <span class="negrita">listas de espera</span>, 
        <span class="negrita">establecimientos sin cuentas activas</span> y otros indicadores claves.
      </p>
      
      <p class="parrafo">
        Entendemos que en este momento los equipos directivos de los establecimientos se encuentran de vacaciones y/o recién retomando sus actividades,
        por lo que compartimos esta información para resguardar la oportuna gestión de las listas de espera.
      </p>
      
      <p class="parrafo">
        <span class="negrita"> La información de cuentas activas corresponde al Miércoles 11 de febrero</span>, 
        última información a la que la DEP tuvo acceso, por lo que es probable que los listados pueden estar desactualizados.
      </p>
      
      <p class="parrafo">
        En este periodo de traspaso, seguiré enviándoles estos reportes, por lo que pueden escribirle a Gabriela Campos (gabriela.campos@dep.cl) y a la jefatura del subdepartamento Javiera Martinez (javiera.martinez@dep.cl) .<br>
        <br>
        Saludos cordiales.
      </p>
      
      <div class="firma">
        <p><strong>Alonso Arraño Portuguez</strong><br>
        Profesional Subdepartamento de Estudios, Monitoreo y Datos <br>
        Subdirección de Desarrollo Estratégico<br>
        Dirección de Educación Pública<br>
        Email: alonso.arrano@dep.cl</p>
      </div>
      
      <div class="nota">
        <strong>Nota:</strong> El archivo debe ser descargado y visualizado desde un navegador, 
        no es compatible con la vista desde el teléfono.
      </div>
    </div>
  </body>
  </html>
  ')
  
  return(html)
}