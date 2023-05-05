# Anteproyecto

## OBJETIVOS

La Aplicación a desarrollar permitirá al usuario cargar una fuente de datos (.exel, csv,...) para manejar los diferentes registros y usarlos como parámetros en distintos tipos de documentos e informes (.odt, .docx, ...) cargados previamente, con el objetivo de automatizar la generación de informes pdf por cada registro.

El proyecto será utilizado en el ámbito de la automatización de informes.

## PREANALISIS DE LO EXISTENTE (Opcional)

*[TODO] Si procede, se informará brevemente sobre el funcionamiento del sistema actual. El que vamos a reemplazar o a mejorar. Este sistema no tiene por qué estar necesariamente automatizado pudiendo realizarse actualmente de forma manual por personas.*

## ANÁLISIS DEL SOFTWARE

En esta sección se explican los requisitos fundamentales que requiere el sistema, es decir, lo que el programa hará desde el punto de vista del cliente.<br>
Dichas pautas y requisitos a implementar en la aplicación son los siguientes:

<ul>
  <li>Permitir al usuario importar una fuente de datos, de forma que esta pueda ser posteriormente usada y gestionada por el programa.</li>
  <li>Permitir seleccionar los archivos o documentos con los que el programa generará los informes de manera automática.</li>
  <li>Automatización en la creación de informes/documentos con los parámetros establecidos por la fuente de datos y los archivos/documentos seleccionados por el usuario previamente en el manejo del programa.</li>
  <li>Trabajar con variables personalizadas para la gestión de los parámetros que la aplicación utilizará a la hora de generar los informes.</li>
  <li>Permitir al usuario manejar todas las diferentes acciones que la aplicación ofrece.</li>
</ul>

### Casos de uso

A continuación se desarrollarán los casos de uso del sistema que capturarán sus requisitos funcionales para expresarlos desde el punto de vista del usuario, los cuales guiarán todo el proceso de desarrollo del sistema.<br>
Estos casos de uso proporcionarán, por tanto, un modo claro y preciso de comunicación entre usuario y desarrollador.<br>
El sistema que se describe en este caso de uso es el siguiente: Un usuario interactúa con el programa y selecciona una fuente de datos y un documento con parámetros introducidos manualmente. Una vez generados los informes, los parámetros serán reemplazados por las variables de la fuente de datos.<br>
Entre las acciones más sencillas y directas que puede realizar están:

<ul>
  <li>Importar documento: Importa un documento cuyos parámetros
introducidos por el usuario serán sustituidos por los valores
que tengan según la fuente de datos.</li>
  <li>Importar fuente de datos: Selecciona un archivo que sirva como
fuente de datos (p. ej. .exel o .csv) para que la aplicación gestione
el valor de los parámetros que se encuentran en el documento actual.</li>
  <li>Generar informes: Selecciona el destino en donde se crearán los
informes pdf acordes con el documento y la fuente de datos proporcionados.
Una vez seleccionado se crearán distintos informes.
  <li>Exportar documentos: </li> Tiene la misma funcionalidad que
<b>generar informes</b>, salvo que el formato de exportación viene a
ser el mismo que el del documento provisto.
</ul>

![Caso de uso de la aplicación](/docs/Caso_de_uso_aplicacion.png)

## DISEÑO DEL SOFTWARE

Las principales librerías que utilizará el software serán las librerías de javafx y Apache POI, la cuál permite gestionar y manejar diferentes tipos de documentos desarrollados por microsoft.

### Ciclo de Vida

En este apartado se explicará el modelo de ciclo de vida elegido para esta aplicación y se explicará la justificación de los procedimientos elegidos para su progresiva implementación desde diferentes puntos de vista.

El modelo de ciclo de vida elegido para esta aplicación es el de <b> Desarrollo en Cascada </b>.

Los motivos por los que he elegido este método de desarrollo para esta aplicación son los siguientes:

<ul>
  <li>Desde el punto de vista del usuario final: El ciclo de vida elegido ofrece la posibilidad de gestionar la aplicación de una manera rápida y sencilla, sin tener que invertir más tiempo del necesario en el proceso gracias a la estructuración de los procesos.</li>
  <li>Desde el punto de vista del Programador: Ofrece una visión sencilla de las ideas a implementar y a cómo debe hacer que el sistema de su programa funcione.</li>
  <li>Por el tipo de aplicación: Permite organizar los diferentes tipos deacciones, interfaces y programas de una manera más eficiente.</li>
  <li>Por la facilidad de uso: Gracias a que el desarrollo en cascada es caracterizado por ordenar de manera rigurosa las etapas del ciclo de vida de software, permite hacerse una idea de cómo puede uno tener todo organizado a la hora de desarrollar el programa en todos sus aspectos.</li>
</ul>

#### Ventajas:

<ol>
  <li value="1"> Permite la departamentalización y control de gestión.</li>
  <li> El horario se establece con los planos normalmente adecuados paracada etapa de desarrollo.</li>
  <li>Este modelo y sus procesos conducen a entregar el proyecto a tiempo.</li>
  <li>Es sencillo y facilita la gestión de proyectos.</li>
  <li>Permite tener bajo control el proyecto.</li>
  <li>Facilita la cantidad de interacción entre equipos que se produce durante el desarrollo.</li>
</ol>

### Diagrama de Flujo de Datos

A continuación se desarrollará los diagramas de flujos de datos de los diferentes procesos y niveles que existen en el entorno de la aplicación.

El diagrama de flujo de datos de este proyecto no solo corresponde a la
aplicación, sino a cómo se gestiona en parte el el documento y la fuente de datos que maneja, ya que ambos aspectos están relacionados directamente. Los diagramas de flujos de datos entonces quedarían de esta manera:

![Diagrama de contexto](/docs/Diagrama1.png)

![Diagrama nivel 1](/docs/Diagrama2.png)

## ESTIMACIÓN DE COSTES

La estimación de costes monetarios de este proyecto está estimada ser de 0 euros. No se dependerá de ninguna API ni servicios que sean de pago para el desarrollo del proyecto.

La duración del desarrollo del proyecto está estimada entre unas 40 y 60 horas.
