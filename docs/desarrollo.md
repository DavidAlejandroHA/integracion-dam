David Alejandro Hernández Alonso 2º DAM A

# Desarrollo

En este documento se documentan todas las tecnologías (herramientas, lenguajes, servicios, ... ) que se utilizan para el desarrollo del proyecyo.

En este documento se elabora una lista de las herramientas, lenguajes y servicios utilizados en el proyecto para entender qué papel tiene cada una de ellas:

## JavaFX

Es una de las principales librerías utilizadas en el proyecto, la cuál ayuda con diferentes herramientas como propiedades observables, bindeos y multitud de herramientas que facilitan la gestión de código.

## Apache commons-text

Apache Commons Text es una librería enfocada en algoritmos que trabajan sobre Strings. Su principal uso en el proyecto está enfocado en escape de caracteres java de Strings usados en el reemplazo de palabras de los documentos a exportar.

## Apache pdfbox

Apache pdfbox es una libería de código abierto de Java, diseñada para trabajar con documentos PDF. Es usada para la creación de ficheros pdf momentáneos que son usados por otras herramientas para forzar el inicio de los servicios de Office para hacer posible la previsualización de documentos.

## JFoenix

JFoenix es una librería de código abierto de Java, que implementa Google Material Design utilizando difetentes componentes de Java.

Es usado por alguno de los componentes de la aplicación, tales como el JFXDrawer (drawer) que se utiliza a la izquierda de la aplicación a modo de menú despegable que contiene diferentes botones que también forman parte de esta librería.

El uso principal de esta libería en el proyecto es el uso de diferentes componentes de esta libería para su uso en la aplicación y junto a una mejor decoración.

## Ikonli

Ikonli (<b>ikonli-javafx</b>) es una libería que proporciona paquetes de iconos que se pueden usar en aplicaciones Java.

Es usada en el uso de diferentes iconos que aparecen en diferentes botones de la aplicación, con el objetivo de aportar un mejor diseño y decoración al programa.

También se usan otras liberías de Ikonli como <b>ikonli-material-pack</b> o <b>ikonli-materialdesign-pack</b>, que son las responsables de utilizar diferentes iconos en la interfaz gráfica de la aplicación.

## Apache POI

Apache POI es una biblioteclibería de Java para leer y escribir formatos de archivo OOXML y binarios de Microsoft Office.

Su uso en la aplicación consiste en la lectura de documentos .xmls y modificación de documentos .docx, pptx y .xlsx, a través de los diferentes métodos y objetos que la libería ofrece.

Existen otras librerías importadas al proyecto que aumentan las funcionalidades y herramientas de Apache POI, tales como <b>poi-scratchpad</b> y <b>poi-ooxml</b>, para hacer uso de funcionalidades extras.

## Jodconverter

JODConverter es una libería de Java que automatiza las conversiones de documentos usando LibreOffice o Apache OpenOffice.

Su uso reside en las conversiones de documentos a ficheros pdf, tanto para la previsualizaciones de los documentos en un visor pdf o para la conversión y exportación de los documentos generados a pdf.

## Pdfviewfx

Es un componente personalizado que permite que una aplicación muestre archivos PDF. El control utiliza el proyecto PDFBox de Apache.

Es usado en la aplicación para ver tanto los documentos importados como el resultado de los documentos exportados.

## Docx4j

Docx4j es una librería de Java de código abierto (ASLv2) para crear y manipular archivos Microsoft Open XML (Word docx, Powerpoint pptx y Excel xlsx).

A pesar de ser una libería enfocada a la creación y manipulación de documentos Microsoft Open XML, su único uso en la aplicación reside en el objeto XPathFactoryUtil, el cuál permite modificar el modelo de implementación de XPath que usará la aplicación para la búsqueda de nodos xml que contengan el texto de las palabras clave importadas en la fuente de datos.

Es usado para modificar el objeto XPathFactory a una versión más reciente de XPath (la 2.0), que es necesario para poder buscar palabras con carácteres especiales en los diferentes nodos de los documentos, y para ello se utiliza el objeto <b>XPathFactoryImpl</b> de la libería <b>Saxon-HE</b>.

## Saxon-HE

Saxon-HE es una libería Saxon-HE de código abierto disponible bajo la Licencia Pública de Mozilla (Mozilla Public License). Proporciona implementaciones de XSLT 2.0, XQuery 1.0 y XPath 2.0 en el nivel básico de conformidad definido por W3C.

Su único propósito y uso en la aplicación es la búsqueda definida de cadenas de caracteres a través de los diferentes nodos de los documentos importados a la aplicación a través del objeto **XPathFactoryImpl**, cuya utilidad servirá para permitir la creación de objetos XPath que usen las funcionalidades implementadas por Saxon-HE para la búsqueda de texto en diferentes nodos.

## Odfdom-java

ODFDOM es un framework de OpenDocument Format (ODF). Su propósito es proporcionar una forma fácil y común de crear, acceder y manipular archivos ODF (.odt, odp, .ods, etc...), sin necesidad de un conocimiento detallado de la especificación ODF.

Al igual que **Apache POI**, es usado en la lectura de ficheros .ods y en la modificación de  documentos ODF, solo que aquí se trata con diferentes tipos de formato y objetos.

## Log4j

Existen varias liberías de Log4j usadas en la aplicación que han sido importadas en el bloque de <dependencyManagement> en el fichero pom.xml del proyecto, con el único objetivo de solucionar unos errores de unas versiones antiguas que estaban siendo utilizadas por otras dependencias del proyecto.
