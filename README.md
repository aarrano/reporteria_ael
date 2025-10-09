# Reporteria proceso anótate en la Lista - DEP

La reporteria del Anótate en la Lista es un proceso secundario de la gestión del Sistema de Admisión Escolar.

Para la generación de este proceso se utilizan dos fuentes de información:

1.  La BBDD consolidada de información del proceso Anótate en la Lista (AEL).
2.  La BBDD de cuentas del AEL por establecimiento

Los paquetes utilizados en este reporte son:

-   tidyverse
-   data.table
-   openxlsx
-   janitor
-   gt
-   knitr
-   quarto
-   ggrepel

## Códigos del proceso

Para generar el reporte se utilizan dos archivos:

1.  `reporte_ael_slep.R` \longrightarrow Archivo que trabaja los dataframes para poder generar el reporte
2.  `reporteria_ael_slep.qmd` \longrightarrow Archivo Quarto que genera el reporte personalizado para cada reporte.
