# Anteproyecto

> NOTA: Incluir diagramas donde proceda (diagramas de clases, casos de uso, entidad relación, ...).

## OBJETIVOS

*[TODO] Se indicará de forma genérica y sin entrar en concreciones el objetivo
que se pretende alcanzar al realizar el proyecto. Se indicará igualmente donde
será utilizado el proyecto obtenido.*

La Aplicación a desarrollar permitirá al usuario cargar una fuente de datos (.exel, csv,...) para manejar los diferentes registros y usarlos como parámetros en distintos tipos de documentos e informes (.odt, .docx, ...) cargados previamente, con el objetivo de automatizar la generación de informes pdf por cada registro.

El proyecto será utilizado en el ámbito de la automatización de informes.

APACHE POI

## PREANALISIS DE LO EXISTENTE (Opcional)

*[TODO] Si procede, se informará brevemente sobre el funcionamiento del sistema actual. El que vamos a reemplazar o a mejorar. Este sistema no tiene por qué estar necesariamente automatizado pudiendo realizarse actualmente de forma manual por personas.*

## ANÁLISIS DEL SOFTWARE

*[TODO] Determinar de forma genérica lo que tiene que hacer el software y cuáles son los requisitos que debe cumplir.*

*Si el proyecto trata sobre la adaptación o ampliación de algún software existente, se deberá aportar información sobre el mismo (documentos electrónicos, direcciones URL, etc.), delimitando claramente cuál será el trabajo que se realizará y que funcionalidad ya está implementada.*


En esta sección se explican los requisitos fundamentales que requiere el sistema, es decir, lo que el programa hará desde el punto de vista del cliente.<br>
Dichas pautas y requisitos a implementar en la aplicación son los siguientes:

<ul>
  <li>Permitir al usuario importar una fuente de datos, de forma que esta pueda ser posteriormente usada y gestionada por el programa.</li>
  <li>Permitir seleccionar los archivos o documentos con los que el programa generará los informes de manera automática.</li>
  <li>Automatización en la creación de informes/documentos con los parámetros establecidos por la fuente de datos y los archivos/documentos seleccionados por el usuario previamente en el manejo del programa.</li>
  <li>Trabajar con variables personalizadas para la gestión de los parámetros que la aplicación utilizará a la hora de generar los informes.</li>
  <li>Permitir al usuario manejar todas las diferentes acciones que la aplicación ofrece.</li>
</ul>

> *Incuir los diagramas necesarios*

### Casos de uso
A continuación se desarrollarán los casos de uso del sistema que capturarán sus requisitos funcionales para expresarlos desde el punto de vista del usuario, los cuales guiarán todo el proceso de desarrollo del sistema.<br>
Estos casos de uso proporcionarán, por tanto, un modo claro y preciso de comunicación entre usuario y desarrollador.<br>
El sistema que se describe en este caso de uso es el siguiente: Un usuario interactúa con el programa y selecciona una fuente de datos y un documento con parámetros introducidos manualmente. Una vez generados los informes, los parámetros serán reemplazados por las variables de la fuente de datos.<br>
Entre las acciones más sencillas y directas que puede realizar están:

<ul>
  <li>Importar documento: Importa un documento cuyos parámetros introducidos por el usuario serán sustituidos por los valores que tengan según la fuente de datos.</li>
  <li>Importar fuente de datos: Selecciona un archivo que sirva como fuente de datos (p. ej. .exel o .csv) para que la aplicación gestione el valor de los parámetros que se encuentran en el documento actual.</li>
  <li>Generar informes: Selecciona el destino en donde se crearán los informes acordes con el documento y la fuente de datos proporcionados. Una vez seleccionado, se crearán los informes.</li>
  <li>Salir de la aplicación: El usuario sale de la aplicación.</li>
</ul>

![Caso de uso de la aplicación](https://github.com/DavidAlejandroHA/integracion-dam/tree/main/docs/Caso_de_uso_aplicacion.png)

## DISEÑO DEL SOFTWARE

*[TODO] Propuesta de posibles opciones de implementación del software que hay que construir, determinar cómo se va a llevar a cabo la implementación.*

>  *Incluir los diagramas necesarios.*

## ESTIMACIÓN DE COSTES

*[TODO] Estimar el coste que representará la creación del proyecto. Esta estimación será temporal y/o económica si procede (costes de contratación de servicios en la nube, por ejemplo).*

La estimación de costes monetarios de este proyecto está estimada ser de 0 euros. No se dependerá de ninguna API ni servicios que sean de pago para el desarrollo del proyecto.

La duración del desarrollo del proyecto está estimada entre unas 40 y 60 horas.
