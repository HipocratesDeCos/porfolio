# Porfolio profesional — Andrés Ragel

## Abrir en Windows
1. Extrae todo el ZIP en una carpeta (no abras el HTML desde dentro del ZIP).
2. Abre `index.html` para ver el porfolio. Para probar la persistencia entre páginas, usa un servidor HTTP local: abre PowerShell en la carpeta extraída y ejecuta `py -m http.server 8000`. Abre `http://localhost:8000/` en Edge o Chrome.
3. Pulsa «Probar caso de riesgo» para cargar el escenario S03 con datos ficticios, o «Abrir aplicación completa».

## Publicar en GitHub Pages
1. Crea un repositorio público nuevo, por ejemplo `porfolio`, desde GitHub. No utilices un repositorio de trabajo de EIOS.
2. Sube el contenido de esta carpeta (index.html, assets/, eios/, .nojekyll y este LEEME), conservando la estructura. No subas el ZIP como único archivo.
3. En Settings → Pages, elige Deploy from a branch, main, / (root) y guarda.
4. Cuando GitHub indique que la publicación está disponible, usa la URL que aparece en Settings → Pages y comprueba todos los enlaces. La URL prevista, si eliges el nombre `porfolio`, sería https://hipocratesdecos.github.io/porfolio/ (no publicada por este paquete).

## Adaptaciones
- Se mantienen los seis archivos de la aplicación original en eios/. Se añade un adaptador local de la interfaz de almacenamiento y un enlace de vuelta al porfolio. No se han modificado las fórmulas del comparador.
- El acceso al caso de riesgo carga S03 si se utilizan los datos ficticios originales. Si se ha activado un libro importado, usa «Importar» y vuelve a los datos originales antes de probar ese caso.
- Almacenamiento local por navegador, no base de datos compartida. No utilizar información empresarial confidencial en la demo pública. Los datos pueden perderse al limpiar el navegador o alcanzar su cuota de almacenamiento. El comparador no es un sistema certificado para producción.
- La importación conserva su dependencia original de SheetJS (CDN cdnjs). Las fuentes externas de la app son opcionales para la visualización; Internet es necesario para cargar SheetJS si no está en caché.
- Data Analytics Mastery se presenta como formación sin atribuir certificación o finalización no confirmadas. No se declaran ahorros monetarios ni niveles de dominio no proporcionados.

## Comprobaciones realizadas
- Ejecución del HTML original adaptado mediante Chromium con contenido local inyectado (la navegación HTTP directa estaba bloqueada en el entorno de pruebas): 216/216 escenarios internos cumplieron las expectativas; S03 cargó 300 unidades con tesorería roja; no se observaron errores JavaScript en esa ejecución.
- Vista de porfolio comprobada a 1200 y 390 píxeles de ancho, sin desbordamiento horizontal.
- Los enlaces de los dos proyectos externos respondían en la consulta realizada.
- No se ha podido comprobar la importación de un Excel completo, la persistencia real entre páginas en un origen HTTP ni el despliegue público. Antes de compartir el sitio, verifica esas funciones en el navegador de destino.
