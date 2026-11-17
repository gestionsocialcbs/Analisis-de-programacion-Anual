# Tablero de Control · Programa Anual de Actividades

Coordinación de Bienestar Social · División de Apoyo Técnico y Mejora de Procesos (DATMP)

Tablero que compara lo programado en el PAA contra lo realizado en inscripciones y asistencias, por año, OOAD, UOPSI, área normativa, tipo y clave de actividad. Tiene dos pestañas:

1. **Avance programado vs. realizado**: indicadores, mapa por OOAD, evolución anual y mensual, áreas normativas, actividades y tablas de detalle.
2. **Comparativo de indicadores**: inscripciones, asistencias y usuarios nominales realizados, por OOAD, UOPSI y actividad.

## Contenido del repositorio

| Archivo | Para qué sirve |
|---|---|
| `index.html` | El tablero. No requiere instalar nada. |
| `datos/paa_datos.js` | Datos consolidados que lee el tablero. |
| `scripts/actualizar_datos.py` | Regenera `datos/paa_datos.js` a partir de los archivos de Excel. |
| `scripts/requirements.txt` | Paquetes de Python que necesita el script. |
| `.nojekyll` | Indica a GitHub Pages que publique los archivos tal cual. |

## Publicar en GitHub Pages

1. En la cuenta **gestionsocialcbs**, crea un repositorio nuevo, por ejemplo `tablero-paa`, de tipo **Public**.
2. Entra al repositorio, elige **Add file › Upload files** y arrastra **todo el contenido** de esta carpeta (incluidas las carpetas `datos` y `scripts` y el archivo `.nojekyll`). Pulsa **Commit changes**.
3. Ve a **Settings › Pages**. En **Build and deployment**, selecciona **Deploy from a branch**, la rama **main** y la carpeta **/ (root)**. Pulsa **Save**.
4. En uno o dos minutos el tablero quedará disponible en:
   `https://gestionsocialcbs.github.io/tablero-paa/`
5. Si quieres que aparezca en el índice de proyectos de la DATMP (`gestionsocialcbs.github.io/DATMP/`), agrega ahí una tarjeta con esa dirección.

> Nota: el archivo `.nojekyll` empieza con punto y en Windows puede quedar oculto. Si no se sube, el tablero funciona igual; solo evita que GitHub procese los archivos.

## Actualizar los datos

Cuando haya un nuevo corte del registro de inscripciones o un nuevo PAA:

1. Coloca en una carpeta los archivos de Excel, uno por año, con nombres como `PAA_2026.xlsx` e `Inscritos_2026.xlsx` (o `Inscritos2026.xlsx`). Los archivos acumulados de varios años se ignoran.
2. Instala los requisitos una sola vez:
   ```
   pip install -r scripts/requirements.txt
   ```
3. Ejecuta el script indicando la carpeta:
   ```
   python scripts/actualizar_datos.py "C:\Users\...\Documents\CBS\DATMP 2026\Bases claude"
   ```
4. Sube a GitHub el archivo `datos/paa_datos.js` que se generó (reemplaza al anterior). El tablero se actualiza solo.

El script detecta automáticamente el último año, el último mes con inscripciones realizadas (por ejemplo, octubre) y la fecha del archivo de inscripciones. Los textos del tablero que mencionan el periodo, como «enero a octubre de 2026», se ajustan con esos datos.

## Reglas de cálculo

- **Programado:** suma de `CAN_CUPO` (inscripciones) y `CAN_ASISTENCIAS` (asistencias) del PAA.
- **Realizado:** suma de `CAN_INSCRITOS` y `CAN_ASISTENCIAS` del registro de inscripciones.
- **Comparación:** por año, OOAD, UOPSI y clave de actividad, sin importar el grupo.
- **Último año:** se compara de enero al mes de corte, tanto lo programado como lo realizado. Los años anteriores se comparan completos.
- **Áreas:** CASSAAM, Atención a Personas con Discapacidad y Servicios de Enlace se integran a Promoción de la Salud; CECART, a Capacitación y Adiestramiento Técnico.
- **Tipo de actividad:** tercer dígito de `CVE_ACTIVIDAD` (1 = curso; 7 y 8 = taller).
- **Asistencias programadas por mes:** se reparten entre los meses que dura cada apertura, en proporción a los días.
- **Semáforo de avance:** morado, más de 125 %; verde, de 85 % a 115 %; amarillo, de 75 % a 85 % o de 115 % a 125 %; rojo, menos de 75 %.

## Pendiente

- **Usuarios nominales:** la pestaña Comparativo ya tiene su espacio. Se integrarán cuando se cuente con sus bases de datos (con año, OOAD, UOPSI y clave de actividad).
