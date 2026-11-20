# Tablero de Control · Programa Anual de Actividades

Coordinación de Bienestar Social · División de Apoyo Técnico y Mejora de Procesos (DATMP)

Tablero que compara lo programado en el PAA contra lo realizado en inscripciones y asistencias, por año, OOAD, UOPSI, CEC, área normativa, tipo y clave de actividad. El tablero está en una sola vista:

1. **Datos generales**: inscripciones, asistencias y usuarios nominales realizados, más las asistencias por inscripción y las inscripciones por usuario nominal.
2. **Avance programado vs. realizado**: indicadores, perfil de la población inscrita (por inscripciones o por usuarios nominales), mapa por OOAD, evolución anual y mensual, áreas normativas y actividades.
3. **Tablas por OOAD y por actividad**, cada una con dos vistas: avance programado vs. realizado y datos reales.

## Exportar

En la barra de filtros hay dos botones:

- **Exportar PDF**: reporte con una portada que detalla los datos descargados (año, OOAD, UOPSI, CEC, área, tipo, grupo y actividad), los apartados 01 a 08 tal como se ven en el tablero y las tablas completas por OOAD y por actividad, con encabezado institucional y numeración de páginas.
- **Exportar Excel**: libro con una hoja de portada (filtros y contenido) y una hoja por apartado, con el formato institucional, totales con fórmulas, filtros de columna y una hoja «Base detallada» para tablas dinámicas. Incluye inscripciones y asistencias, sin importar la unidad elegida en el tablero.

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

1. Coloca en una carpeta los archivos de Excel, uno por año, con nombres como `PAA_2026.xlsx` e `Inscritos_2026.xlsx` (o `Inscritos2026.xlsx`), junto con `Catálogo agrupado.xlsx`. Los archivos acumulados de varios años se ignoran.
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

- **Grupo de actividad:** se toma de la hoja «Agrupaciones» del archivo `Catálogo agrupado.xlsx` (columna Agrupación), por año y clave de actividad. Si una clave no aparece en el año, se usa su agrupación más frecuente; las claves 31806 (Verano entre telones) y 43701 (Curso de habilidad física, club), que no están en el catálogo, se asignan a Curso Vacacional y a Personas mayores y cuidadores, respectivamente; cualquier otra clave ausente se marca como «Sin agrupación». Al elegir un grupo, el filtro de actividad solo muestra las actividades de ese grupo.
- **Usuarios nominales** (datos generales, tablas de indicadores realizados y perfil): se toman de la carpeta `Nominales` (un archivo por año, de 2021 a 2025; todas las hojas con `REF_CURP`). Se cuentan personas únicas por `REF_CURP`. Si el CURP está vacío, se busca el de la misma persona por clave de alumno o por nombre completo y fecha de nacimiento; si no existe, se le asigna un identificador propio para que la inscripción cuente. Al sumar varios años, una persona se cuenta una vez en cada año; en los desgloses por OOAD, UOPSI o actividad, una persona se cuenta en cada renglón en el que tiene inscripciones. El perfil de usuarios nominales clasifica a cada persona con el sexo, la edad (columna `Edad_calculada`; si no existe, `EDAD` o `CAN_EDAD`) y la calidad jurídica (`NOM_CONDICION` y `NOM_TIPO_USUARIO`) que más se repiten en sus inscripciones del año. El tablero solo publica totales, nunca datos personales.
- **CEC:** se toma de la columna `CVE_CEC` del PAA y del registro de inscripciones: 0 corresponde a UOPSI y cualquier valor mayor a 0, a CEC.
- **Filtros:** año, OOAD, UOPSI, CEC, área normativa, tipo, grupo de actividad y actividad admiten una, varias o todas las opciones. Cada lista muestra solo las opciones con datos para el resto de los filtros.

- **Programado:** suma de `CAN_CUPO` (inscripciones) y `CAN_ASISTENCIAS` (asistencias) del PAA.
- **Realizado:** suma de `CAN_INSCRITOS` y `CAN_ASISTENCIAS` del registro de inscripciones.
- **Comparación:** por año, OOAD, UOPSI y clave de actividad, sin importar el grupo.
- **Último año:** se compara de enero al mes de corte, tanto lo programado como lo realizado. Los años anteriores se comparan completos.
- **Áreas:** CASSAAM, Atención a Personas con Discapacidad y Servicios de Enlace se integran a Promoción de la Salud; CECART, a Capacitación y Adiestramiento Técnico.
- **Tipo de actividad:** tercer dígito de `CVE_ACTIVIDAD` (1 = curso; 7 y 8 = taller).
- **Asistencias programadas por mes:** se reparten entre los meses que dura cada apertura, en proporción a los días.
- **Perfil de la población inscrita** (solo inscripciones realizadas): género (`CAN_MASCULINOS`, `CAN_FEMENINOS`); grupo etario: jóvenes (`CAN_00_A_02` a `CAN_16_A_19`), adultos (`CAN_20_A_59`) y adultos mayores (`CAN_MAS_DE_60`); y calidad jurídica (las 16 columnas `CAN_DERECHOHABIENTE` a `CAN_ADULTO_MAYOR_NODERER`). Solo se muestran las categorías con registros.
- **Semáforo de avance:** morado, más de 125 %; verde, de 85 % a 115 %; amarillo, de 75 % a 85 % o de 115 % a 125 %; rojo, menos de 75 %.
