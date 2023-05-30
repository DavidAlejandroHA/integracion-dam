package integracion.wordseedexporter.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;

import org.apache.commons.io.FileUtils;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFEndnote;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFFootnote;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.docx4j.utils.XPathFactoryUtil;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.JodConverter;
import org.odftoolkit.odfdom.doc.OdfGraphicsDocument;
import org.odftoolkit.odfdom.doc.OdfPresentationDocument;
import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.doc.presentation.OdfSlide;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.dom.element.draw.DrawPageElement;
import org.odftoolkit.odfdom.dom.element.office.OfficeDrawingElement;
import org.odftoolkit.odfdom.dom.element.office.OfficeSpreadsheetElement;
import org.odftoolkit.odfdom.dom.element.office.OfficeTextElement;
import org.w3c.dom.NodeList;

import integracion.wordseedexporter.controllers.Controller;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import net.sf.saxon.xpath.XPathFactoryImpl;

/**
 * <p>
 * Esta clase se encarga de, dado un documento entregado, generar distintos<br>
 * documentos, cada uno con los cambios de palabras correspondientes<br>
 * respecto a la fuente de datos previamente importada en la aplicación.
 * </p>
 * <p>
 * Las palabras reemplazadas en el documento dependen además de la<br>
 * BooleanProperty {@link BooleanProperty replaceExactWord} de la clase
 * Controller<br>
 * (Controller.replaceExactWord), y de {@link BooleanProperty caseSensitive}
 * (Controller.caseSensitive) <br>
 * ya que según su valor llegará a reemplazar solo las palabras exactas <br>
 * que el usuario haya especificado (true) o cualquier coincidencia<br>
 * que incluyan las que estén dentro de otra palabra que se encuentre en el<br>
 * documento (false). <br>
 * </p>
 * 
 * @author David Alejandro
 *
 */
public class DocumentManager {

	public DocumentManager() {
	}

	/**
	 * Este método gestiona el tipo de documento entregado (.docx, pptx, xlsx, odt,
	 * odp, ods o odg) y ejecuta el método apropiado para editar el documento.
	 * 
	 * @param input  El fichero (documento) a modificar
	 * @param output La ruta en la que los documentos generados se almacenarán
	 * @throws Exception
	 */
	public void giveDocument(File input, File output, boolean exportarAPDF) throws Exception {

		if (input != null) {
			ObservableList<DataSource> dataSources = Controller.dataSources;

			if (dataSources != null && dataSources.size() > 0) {
				XPathFactoryUtil.setxPathFactory(new XPathFactoryImpl());
				if (input.getName().endsWith(".docx")) {
					replaceDocxStrings(dataSources, input, output, exportarAPDF);

				} else if (input.getName().endsWith(".pptx")) {
					replacePptxStrings(dataSources, input, output, exportarAPDF);

				} else if (input.getName().endsWith(".xlsx")) {
					replaceXlsxStrings(dataSources, input, output, exportarAPDF);

				} else {

					if (input.getName().endsWith(".odt")) {
						replaceOdtStrings(dataSources, input, output, exportarAPDF);

					} else if (input.getName().endsWith(".odp")) {
						replaceOdpStrings(dataSources, input, output, exportarAPDF);

					} else if (input.getName().endsWith(".ods")) {
						replaceOdsStrings(dataSources, input, output, exportarAPDF);

					} else if (input.getName().endsWith(".odg")) {
						replaceOdgStrings(dataSources, input, output, exportarAPDF);

					}
				}
			}
		}

	}

	/**
	 * Dependiendo de la extensión del fichero ofrecido que contiene la fuente de
	 * datos necesaria para el programa, se ejecuta el método adecuado para su
	 * correcta lectura de datos
	 * 
	 * @param f El fichero / fuente de datos a importar en la aplicación
	 * @throws Exception
	 */
	public void readData(File f) throws Exception {
		if (f != null) {
			if (f.getName().endsWith(".ods")) {
				readOds(f);
			}
			if (f.getName().endsWith(".xlsx")) {
				readXlsx(f);
			}
		}
	}

	// La lista de listas de string vendrá de la próxima interfaz a crear
	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".docx" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void replaceDocxStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			// La interfaz IBody es la que implementa el método .getParagraphs() y las
			// clases que manejan el contenido de los párrafos en los docx
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);
				XWPFDocument doc = new XWPFDocument(new FileInputStream(input)); // para reiniciar el doc a su estado
																					// inicial
																					// para poder volver a reemplazar
																					// texto

				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;

						List<XWPFParagraph> parrafos = doc.getParagraphs();
						List<XWPFTable> tablas = doc.getTables();
						XWPFComment[] comentarios = doc.getComments();
						List<XWPFEndnote> endNotes = doc.getEndnotes();
						List<XWPFFooter> piesPag = doc.getFooterList();
						List<XWPFFootnote> notasPiesPag = doc.getFootnotes();
						List<XWPFHeader> cabeceras = doc.getHeaderList();

						// Párrafos del cuerpo del documento
						if (parrafos != null) {
							for (XWPFParagraph p : parrafos) {
								cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
								if (cambios) {
									numCambios++;
								}
							}
						}

						// Tablas
						if (tablas != null) {
							for (XWPFTable tbl : tablas) {
								for (XWPFTableRow row : tbl.getRows()) {
									for (XWPFTableCell cell : row.getTableCells()) {
										for (XWPFParagraph p : cell.getParagraphs()) {
											cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
											if (cambios) {
												numCambios++;
											}
										}
									}
								}
							}
						}

						// Comentarios
						if (comentarios != null) {
							for (XWPFComment comms : comentarios) {
								for (XWPFParagraph p : comms.getParagraphs()) {
									cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

						// EndNotes
						if (endNotes != null) {
							for (XWPFEndnote endNote : endNotes) {
								for (XWPFParagraph p : endNote.getParagraphs()) {
									cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

						// Pies de páginas
						if (piesPag != null) {
							for (XWPFFooter footer : piesPag) {
								for (XWPFParagraph p : footer.getParagraphs()) {
									cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

						// Notas de pies de páginas
						if (notasPiesPag != null) {
							for (XWPFFootnote footNote : notasPiesPag) {
								for (XWPFParagraph p : footNote.getParagraphs()) {
									cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

						// Cabeceras
						if (cabeceras != null) {
							for (XWPFHeader header : cabeceras) {
								for (XWPFParagraph p : header.getParagraphs()) {
									cambios = editDocxParagraph(p, cellString.get(j), columnKeyName.get(j));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}
						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve
																							// a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}
				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".docx";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					doc.write(fOut);
					fOut.close();
					createdFiles.add(file);
				}
				doc.close();
			}
		}

		// Eliminar los posibles ficheros que pueden haber quedado de acciones pasadas,
		// sin incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".docx");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".pptx" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replacePptxStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);
				XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(input));
				// para reiniciar el doc a su estado inicial para poder volver a reemplazar
				// texto
				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;
						List<XSLFSlide> slides = slideShow.getSlides();
						if (slides != null) {
							for (XSLFSlide slide : slideShow.getSlides()) {
								for (XSLFShape sh : slide.getShapes()) {
									if (sh instanceof XSLFTextShape) {

										List<XSLFTextParagraph> parrafos = ((XSLFTextShape) sh).getTextParagraphs();

										if (parrafos != null) {
											for (XSLFTextParagraph p : parrafos) {
												cambios = editPptxParagraph(p, cellString.get(j), columnKeyName.get(j));
												if (cambios) {
													numCambios++;
												}
											}
										}
									}
								}
							}
						}

						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}
						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve
																							// a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".pptx";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					slideShow.write(fOut);
					fOut.close();
					createdFiles.add(file);
				}
				slideShow.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".pptx");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".xlsx" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replaceXlsxStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);

				XSSFWorkbook spreadSheet = new XSSFWorkbook(new FileInputStream(input));
				// para reiniciar el doc a su estado inicial para poder volver a reemplazar
				// texto
				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;
						Iterator<Sheet> sheets = spreadSheet.sheetIterator();
						if (sheets != null) {
							// Para cada página del documento excel
							while (sheets.hasNext()) {
								Sheet sh = sheets.next(); // Se maneja cada hoja
								for (Row row : sh) {
									for (Cell cell : row) {
										cambios = editXlxsCells(cell, cellString.get(j), columnKeyName.get(j));
										if (cambios) {
											numCambios++;
										}
									}
								}
							}
						}
						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}
						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".xlsx";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					spreadSheet.write(fOut);
					fOut.close();
					createdFiles.add(file);
				}
				spreadSheet.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".xlsx");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".odt" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replaceOdtStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);

				OdfTextDocument odtDocument = OdfTextDocument.loadDocument(input);
				// para reiniciar el doc a su estado inicial para poder volver a reemplazar
				// texto
				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;

						// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
						// En este caso solo se escapea las comillas dobles porque es lo que se usa para
						// el contains"" del xpath
						String escapedKey1 = columnKeyName.get(j);
						escapedKey1 = escapedKey1.replaceAll("\"", "\"\"");

						// Recibiendo el nombre de la palabra clave
						String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

						OfficeTextElement odtTextEl = odtDocument.getContentRoot();
						NodeList odtNodes = odtTextEl.getChildNodes();

						XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
						NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
								XPathConstants.NODESET);
						cambios = editStringNodeList(nodelist, cellString.get(j), columnKeyName.get(j));
						if (cambios) {
							numCambios++;
						}
						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}
						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".odt";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					odtDocument.save(fOut);
					fOut.close();
					createdFiles.add(file);
				}
				odtDocument.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".odt");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".odp" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replaceOdpStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);

				OdfPresentationDocument odpDocument = OdfPresentationDocument.loadDocument(input);
				// para reiniciar el doc a su estado inicial para poder volver a reemplazar
				// texto
				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;

						// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
						// En este caso solo se escapea las comillas dobles porque es lo que se usa para
						// el contains"" del xpath
						String escapedKey1 = columnKeyName.get(j);
						escapedKey1 = escapedKey1.replaceAll("\"", "\"\"");

						// Recibiendo el nombre de la palabra clave
						String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

						Iterator<OdfSlide> it = odpDocument.getSlides();
						while (it.hasNext()) {
							OdfSlide odfSlide = it.next();
							DrawPageElement slideElement = odfSlide.getOdfElement();
							NodeList slideNodeList = slideElement.getChildNodes();

							XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
							NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(slideNodeList,
									XPathConstants.NODESET);
							cambios = editStringNodeList(nodelist, cellString.get(j), columnKeyName.get(j));
							if (cambios) {
								numCambios++;
							}
						}
						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}
						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) {
							notFound.add(columnKeyName.get(j));
						} // si ya lo contiene no se vuelve a agregar
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".odp";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					odpDocument.save(fOut);
					fOut.close();
					createdFiles.add(file);

				}
				odpDocument.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".odp");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".ods" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replaceOdsStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);

				OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(input);
				// para reiniciar el doc a su estado inicial para poder volver a reemplazar
				// texto
				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;

						// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
						// En este caso solo se escapea las comillas dobles porque es lo que se usa para
						// el contains"" del xpath
						String escapedKey1 = columnKeyName.get(j);
						escapedKey1 = escapedKey1.replaceAll("\"", "\"\"");

						// Recibiendo el nombre de la palabra clave
						String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";
						OfficeSpreadsheetElement odtSShEl = odsDocument.getContentRoot();
						NodeList odtNodes = odtSShEl.getChildNodes();

						XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
						NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
								XPathConstants.NODESET);
						cambios = editStringNodeList(nodelist, cellString.get(j), columnKeyName.get(j));
						if (cambios) {
							numCambios++;
						}

						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}

						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".ods";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					odsDocument.save(fOut);
					fOut.close();
					createdFiles.add(file);

				}
				odsDocument.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".ods");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Edita el fichero documento entregado previamente por el método
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de importación de la fuente de datos.
	 * 
	 * @param dataSources Las fuentes de datos a procesar para reemplazar las
	 *                    palabras en documentos
	 * @param input       El documento importado
	 * @param output      La ruta del directorio donde se generarán los
	 *                    documentos/pdf
	 * @param pdf         Si se exportarán los archivos en ".odg" o en ".pdf"
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	private void replaceOdgStrings(ObservableList<DataSource> dataSources, File input, File output, boolean pdf)
			throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<File> createdFiles = new ArrayList<>();

		for (int dim = 0; dim < dataSources.size(); dim++) {
			ObservableList<String> columnKeyName = dataSources.get(dim).getKeyNames();
			ObservableList<ObservableList<String>> rows = dataSources.get(dim).getRows();

			iEffective = 0;
			jEffective++;
			for (int i = 0; i < rows.size(); i++) { // iterando en las filas
				numCambiosRow = 0;

				List<String> cellString = rows.get(i);

				OdfGraphicsDocument odgDocument = OdfGraphicsDocument.loadDocument(input); // para reiniciar el doc a su
																							// estado inicial para poder
																							// volver a reemplazar texto

				for (int j = 0; j < columnKeyName.size(); j++) {// iterando en los elementos de cada fila junto a su
																// correspondiente palabra clave
					if ((columnKeyName.get(j) != null && columnKeyName.get(j).trim().length() > 0)
							&& (cellString.get(j) != null && cellString.get(j).trim().length() > 0)) {

						numCambios = 0;
						boolean cambios = false;

						// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
						// En este caso solo se escapea las comillas dobles porque es lo que se usa para
						// el contains"" del xpath
						String escapedKey1 = columnKeyName.get(j);
						escapedKey1 = escapedKey1.replaceAll("\"", "\"\"");

						// Recibiendo el nombre de la palabra clave
						String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";
						OfficeDrawingElement odgDrawEl = odgDocument.getContentRoot();
						NodeList odtNodes = odgDrawEl.getChildNodes();

						XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
						NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
								XPathConstants.NODESET);
						cambios = editStringNodeList(nodelist, cellString.get(j), columnKeyName.get(j));
						if (cambios) {
							numCambios++;
						}

						if (numCambios > 0) { // si ha habido cambios
							numCambiosRow++;
						}

						if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se
																							// vuelve a agregar
							notFound.add(columnKeyName.get(j));
						}
					}
				}

				if (numCambiosRow > 0) {
					iEffective++;
					String fileName = "output_" + (jEffective) + "_" + (iEffective) + ".odg";
					File file = new File(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName);
					FileOutputStream fOut = new FileOutputStream(file);
					odgDocument.save(fOut);
					fOut.close();
					createdFiles.add(file);
				}
				odgDocument.close();
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);
		if (Controller.officeInstalled) {
			convertFileListToPdfs(createdFiles, ".odg");
		}
		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
		moveFiles(createdFiles, output, pdf);
	}

	/**
	 * Mueve ficheros de la carpeta tmpDocs de la aplicación hacia la ruta
	 * especificada.
	 * 
	 * @param createdFiles La lista de ficheros a mover a otro destino
	 * @param output       La carpeta/dirección a la que se moverán los ficheros
	 * @param pdf          Si se moverán solo los documentos generados o se copiarán
	 *                     las previsualizaciones pdf a la carpeta de destino
	 * @throws IOException
	 */
	private void moveFiles(List<File> createdFiles, File output, boolean pdf) throws IOException {
		if (!pdf) {
			for (int i = 0; i < createdFiles.size(); i++) {
				try {
					Files.move(Paths.get(createdFiles.get(i).getPath()),
							Paths.get(output.getPath() + File.separator + createdFiles.get(i).getName()),
							StandardCopyOption.REPLACE_EXISTING);
				} catch (Exception e) {
				}
			}
		} else {
			for (File f : Controller.previsualizaciones) {
				try { // Copiando las previsualizaciones pdf a la ruta de exportación
					Files.copy(Paths.get(f.getPath()), Paths.get(output.getPath() + File.separator + f.getName()),
							StandardCopyOption.REPLACE_EXISTING);
				} catch (Exception e) {
				}
			}

			// borrando los documentos generados que servían para pasarlos a pdf
			for (int i = 0; i < createdFiles.size(); i++) {
				try {
					Files.delete(Paths.get(createdFiles.get(i).getPath()));
				} catch (Exception e) {
				}
			}
		}
	}

	/**
	 * Edita el párrafo de un archivo docx proporcionado al método de forma que si
	 * encuentra la palabra a reemplazar que se le ha dado al método, la reemplazará
	 * en ese párrafo al nuevo valor adjuntado en este mismo método
	 * 
	 * @param p      El párrafo a editar
	 * @param newKey El nuevo valor que tendrá la palabra clave a reemplazar
	 * @param key    La palabra clave a reemplazar
	 * @return true Si se han hecho cambios en el párrafo, false si no ha sido así
	 */
	private boolean editDocxParagraph(XWPFParagraph p, String newKey, String key) {
		String regexKey = stringModifyOptions(key);
		boolean cambios = false;
		int numCambios = 0;
		List<XWPFRun> runs = p.getRuns();
		if (runs != null && !newKey.equals(key)) { // si las dos claves son iguales de nada sirve hacer cambios
			for (XWPFRun r : runs) {
				String text = r.getText(0);
				String textAux = text;
				if (text != null && text.contains(key)) {
					text = text.replaceAll(regexKey, StringEscapeUtils.escapeJava(newKey));
					r.setText(text, 0);
					if (!text.equals(textAux)) { // si se cambió algo del texto entonces se hicieron cambios
						numCambios++;
					}
				}
			}
		}
		if (numCambios > 0) {
			cambios = true;
		}

		return cambios;
	}

	/**
	 * Edita el párrafo de un archivo pptx proporcionado al método de forma que si
	 * encuentra la palabra a reemplazar que se le ha dado al método, la reemplazará
	 * en ese párrafo al nuevo valor adjuntado a este método
	 * 
	 * @param p      El párrafo a editar
	 * @param newKey El nuevo valor que tendrá la palabra clave a reemplazar
	 * @param key    La palabra clave a reemplazar
	 * @return true Si se han hecho cambios en el párrafo, false si no ha sido así
	 */
	private boolean editPptxParagraph(XSLFTextParagraph p, String newKey, String key) {
		String regexKey = stringModifyOptions(key);

		boolean cambios = false;
		int numCambios = 0;
		List<XSLFTextRun> runs = p.getTextRuns();
		if (runs != null) {
			for (XSLFTextRun r : runs) {
				String text = r.getRawText();
				String textAux = text;
				if (text != null && text.contains(key)) {
					text = text.replaceAll(regexKey, StringEscapeUtils.escapeJava(newKey));
					r.setText(text);
					if (!text.equals(textAux)) { // si se cambió algo del texto entonces se hicieron cambios
						numCambios++;
					}
				}
			}
		}
		if (numCambios > 0) {
			cambios = true;
		}

		return cambios;
	}

	/**
	 * Edita la celda proporcionada a este método de un archivo .xlxs de forma que
	 * si encuentra la palabra a reemplazar que se le ha dado a este método, la
	 * reemplazará en esa celda al nuevo valor adjuntado a este método
	 * 
	 * @param c
	 * @param newKey
	 * @param key
	 * @return true Si se han hecho cambios en la celda, false si no ha sido así
	 */
	private boolean editXlxsCells(Cell c, String newKey, String key) {
		String regexKey = stringModifyOptions(key);

		boolean cambios = false;
		int numCambios = 0;
		// Solo modificar las celdas que contengan contenido de texto
		// Si contienen otras cosas que no sean texto, se ignoran ya que al modificar
		// por ejemplo celdas numéricas con texto, podría lanzarse un error
		if (c.getCellType().equals(CellType.STRING)) {
			String text = c.getStringCellValue();
			String textAux = text;
			if (text != null && text.contains(key)) {
				text = text.replaceAll(regexKey, StringEscapeUtils.escapeJava(newKey));
				c.setCellValue(text);
				if (!text.equals(textAux)) { // si se cambió algo del texto entonces se hicieron cambios
					numCambios++;
				}
			}
		}
		if (numCambios > 0) {
			cambios = true;
		}

		return cambios;
	}

	/**
	 * Edita el texto de los distintos nodos ubicados en la lista de nodos
	 * proporcionada a este método, de forma que si encuentra la palabra a
	 * reemplazar que se le ha dado a este método, la reemplazará en ese texto al
	 * nuevo valor adjuntado al método
	 * 
	 * @param c
	 * @param newKey
	 * @param key
	 * @return true Si se han hecho cambios de texto dentro del nodo, false si no ha
	 *         sido así
	 */
	private boolean editStringNodeList(NodeList nl, String newKey, String key) {
		boolean cambios = false;
		int numCambios = 0;
		if (nl != null) {
			for (int k = 0; k < nl.getLength(); k++) {
				if (nl.item(k).getTextContent().contains(key)) {
					String texto = nl.item(k).getTextContent();
					String textAux = texto;
					String controlRegex = stringModifyOptions(key);
					texto = texto.replaceAll(controlRegex, StringEscapeUtils.escapeJava(newKey));
					nl.item(k).setTextContent(texto);
					if (!texto.equals(textAux)) { // si se cambió algo del texto entonces se hicieron cambios
						numCambios++;
					}
				}
			}
		}
		if (numCambios > 0) {
			cambios = true;
		}

		return cambios;
	}

	/**
	 * Este método se encarga de añadirle los carácteres necesarios para que
	 * posteriormente pueda ser utilizada en los métodos .replaceAll() del remplazo
	 * de texto en los distintos párrafos y nodos de los documentos de la manera
	 * especificada por el usuario a través de las opciones de exportación de la
	 * aplicación.
	 * 
	 * @param k
	 * @return
	 */
	private String stringModifyOptions(String k) {
		k = Pattern.quote(k); // con este método se obtiene un patrón "literal" del string, deforma que los
								// caracteres especiales que puedatener el string quedan "asegurados" en lo que
								// se refiere a la sintaxis de las expresiones regulares que utiliza el
								// método .replaceAll(), evitando así posible errores de sintaxis en
								// estas expresiones
								// Ojo: no es lo mismo que Matcher.quoteReplacement(k)
		if (!Controller.caseSensitive.get()) {
			// Si la BooleanProperty replaceExactWord está a true, se le añade a k
			// una expresión regular (?i) para que el replaceAll() acepte tanto mayúsculas
			// como minúsculas (case insensitive)
			k = "(?i)" + k;
		}
		if (Controller.replaceExactWord.get()) {
			// Si la BooleanProperty replaceExactWord está a true, se le añaden a k un
			// "limitador" de palabras para seleccionar solo esas mismas palabras (\\b + k +
			// \\b) y no otras que la puedan contener pero que no sean la misma, ya que el
			// .replaceAll() puede trabajar con expresiones regulares (regex)
			// Al convertir la string a una string literal (Pattern.quote(k)), se puede
			// tratar la string entera como una palabra y el limitador de palabras \\b no
			// daría problemas con ningún carácter
			k = "\\b" + k + "\\b";
		}
		return k;
	}

	/**
	 * Crea una alerta advirtiendo de que varias palabras especificadas en la fuente
	 * de datos no se han encontrado en el documento importado, por lo que varios
	 * ficheros no contienen todos los cambios de la fuente de datos
	 * 
	 * @param ls La lista de palabras sin encontrar que saldrán en la alerta
	 */
	private void notFoundWords(List<String> ls) {
		Alert alerta = new Alert(AlertType.WARNING);
		alerta.setTitle("No se han encontrado algunas palabras");
		String textoPalabras = "";
		for (int i = 0; i < ls.size(); i++) {
			textoPalabras += ((i < ls.size() - 1) ? "" : "y ") + "\"" + ls.get(i) + "\""
					+ ((i < ls.size() - 1) ? ", " : " ");
		}
		Controller.crearAlerta(AlertType.WARNING, "No se han encontrado algunas palabras",
				"Las palabras " + textoPalabras + "no se han encontrado en el fichero importado.",
				"Varias palabras especificadas en la fuente de datos no han sido encontradas, por lo que no\n"
						+ "se han aplicado los cambios correspondientes respecto a estas palabras.",
				false);
	}

	/**
	 * Crea una alerta informado de que todas las palabras de la fuente de datos se
	 * encontraron en el documento importado, por lo que se realiza el reemplazo de
	 * las palabras de la fuente de datos.
	 */
	private void allWordsSuccess() {
		Controller.crearAlerta(AlertType.INFORMATION, "Operación exitosa",
				"Se han reemplazado todas las palabras de la fuente de datos\n"
						+ "y aplicado los cambios correspondientes.",
				null, false);
	}

	/**
	 * Elimina los posibles ficheros viejos (documentos) que queden almacenados en
	 * la carpeta tmpDocs de la aplicación.
	 * 
	 * @param ls
	 */
	private void deleteOldFiles(List<File> ls) {
		List<File> fileList = new ArrayList<File>(
				FileUtils.listFiles(new File(Controller.TEMPDOCSFOLDER.getPath()), null, false));
		for (File file : fileList) {
			if (!ls.contains(file) && !file.getName().endsWith(".pdf")) {
				FileUtils.deleteQuietly(file);
			}
		}
	}

	/**
	 * El método se encarga de leer los archivos ".ods" y transformar sus datos en
	 * una fuente de datos para el programa
	 * 
	 * @param f El fichero a leer los datos
	 * @throws Exception
	 */
	private void readOds(File f) throws Exception {
		OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(f);
		List<OdfTable> tablas = odsDocument.getTableList(false);
		if (tablas != null) {

			ObservableList<DataSource> dsList = FXCollections.observableArrayList();

			for (OdfTable t : tablas) {
				// leer el tamaño de la tabla
				int width = 0;
				int height = 0;
				int rowIndexStart = 0;
				int columnIndexStart = 0;
				boolean lock = false;

				DataSource ds = new DataSource();

				ObservableList<String> celdas = FXCollections.observableArrayList();
				ObservableList<String> nombresReemplazo = FXCollections.observableArrayList();

				ObservableList<ObservableList<String>> filas = FXCollections
						.<ObservableList<String>>observableArrayList();

				// se obtiene un tamaño para buscar los registros de la tabla y desde donde
				// empieza
				int searchHeight = t.getRowElementList().size();
				int searchWidth = 0;

				// encontrando la anchura máxima
				for (int i = 0; i < searchHeight; i++) {
					if (t.getRowElementList().get(i).getLength() > searchWidth) {
						searchWidth = t.getRowElementList().get(i).getLength() + 1;
					}
				}

				// lecutra de las dimensiones reales (quitando las columnas en blanco) de la
				// tabla
				// y su punto de partida
				for (int i = 0; i < searchHeight; i++) {
					boolean contains = false;

					celdas = FXCollections.observableArrayList(); // reset de filas
					for (int j = 0; j < searchWidth; j++) {
						OdfTableCell cell = t.getCellByPosition(j, i);
//						if (emptyCellsCount > 100) {
//							throw new Exception("La tabla contiene demasiados registros vacíos. Es posible que\n"
//									+ "sea necesario cortar y pegar la tabla en el inicio de un nuevo documento.");
//						}
						if (cell != null && cell.getValueType() != null && cell.getStringValue().trim().length() > 0) {
							contains = true;
							if (!lock) {
								rowIndexStart = i;
								columnIndexStart = j;
								lock = true;
							}

							if (rowIndexStart == i) { // si está en la primera fila y lee un registro
								width = j - columnIndexStart + 1; // apunta la anchura hasta llegar a el último de ellos
							}
						}
					}
					if (contains) {
						height++;
					}
				}

				if (height <= 1) {
					throw new Exception("Las tablas deben de tener más de una fila.");
				}

				String texto = null;
				for (int i = rowIndexStart; i < rowIndexStart + height; i++) {
					celdas = FXCollections.observableArrayList(); // reset de celdas
					for (int j = columnIndexStart; j < columnIndexStart + width; j++) {
						OdfTableCell cell = t.getCellByPosition(j, i);
						if (cell == null) { // si hay datos el texto son los datos de la casilla
							texto = ""; // si no se pone a string vacío
						} else {
							texto = cell.getStringValue();
						}
						if (i == rowIndexStart) { // si es el primer registro se usa como una palabra a reemplazar en el
							// documento
							nombresReemplazo.add(texto);
						} else { // si no como una de las claves a usar para el reemplazo de palabras
							celdas.add(texto);
						}
					}
					if (celdas.size() > 0) { // celdas = 1 fila
						filas.add(celdas); // una vez procesadas todas las filas se añaden a la lista de filas
					}
				}

				// eliminar las columnas que no tienen nombres clave
				for (int i = 0; i < nombresReemplazo.size(); i++) {
					if (nombresReemplazo.get(i).trim().length() == 0) {
						nombresReemplazo.remove(i);
						for (int j = 0; j < filas.size(); j++) {
							filas.get(j).remove(i);
						}
						i--;
					}
				}
				ds.setKeyNames(nombresReemplazo);
				ds.setRows(filas);
				dsList.add(ds);
			}
			Controller.dataSources.set(dsList);
		}
	}

	/**
	 * El método se encarga de leer los archivos ".xlsx" y transformar sus datos en
	 * una fuente de datos para el programa
	 * 
	 * @param f El fichero a leer los datos
	 * @throws Exception
	 */
	private void readXlsx(File f) throws Exception {
		XSSFWorkbook spreadSheet = new XSSFWorkbook(new FileInputStream(f));

		Iterator<Sheet> sheets = spreadSheet.sheetIterator();

		if (sheets != null) {
			ObservableList<DataSource> dSList = FXCollections.observableArrayList();

			// Para cada página del documento excel
			while (sheets.hasNext()) {
				DataSource ds = new DataSource();

				ObservableList<String> rowElements = FXCollections.observableArrayList();
				ObservableList<String> nombresReemplazo = FXCollections.observableArrayList();

				ObservableList<ObservableList<String>> rowList = FXCollections.observableArrayList();

				Sheet sh = sheets.next(); // Se maneja cada hoja

				// leer el tamaño de la tabla
				int width = 0;
				int height = 0;
				int rowIndexStart = 0;
				int columnIndexStart = 0;
				boolean lock = false;
				for (int i = sh.getFirstRowNum(); i < sh.getLastRowNum() + 1; i++) {
					Row row = sh.getRow(i);
					boolean contains = false;
					if (row != null) { // si la celda no es nula
						int firstCellAux = row.getFirstCellNum();
						int lastCellAux = row.getLastCellNum();

						for (int colNum = firstCellAux; colNum < lastCellAux; colNum++) {
							Cell cell = row.getCell(colNum);
							if (cell != null && readCell(cell).trim().length() > 0) { // si hay algo en la celda
								contains = true;
								if (!lock) {
									rowIndexStart = cell.getRowIndex();
									columnIndexStart = cell.getColumnIndex();
									width = lastCellAux - firstCellAux; // el ancho de la tabla vendrá
									// dado por el ancho de la primera fila
									lock = true;
								}
							}
						}
						if (contains) {
							height++;
						}
					}
				}
				if (height <= 1) {
					throw new Exception("Las tablas deben de tener más de una fila.");
				}

				for (int i = rowIndexStart; i < rowIndexStart + height; i++) { // manejando cada fila
					rowElements = FXCollections.observableArrayList();
					Row row = sh.getRow(i);

					if (row != null) {
						for (int cn = columnIndexStart; cn < columnIndexStart + width; cn++) { // manejando
							// cada celda de las filas
							Cell cell = row.getCell(cn);
							String texto = "";
							if (cell != null) {
								texto = readCell(cell); // se empieza a analizar el texto
								int lineaActual = cell.getRowIndex();
								if (lineaActual == rowIndexStart) {
									// si la celda es de los nombres clave
									nombresReemplazo.add(texto); // para añadirlo a la lista de nombres a
									// reemplazar

								} else { // si no se añade a la lista de cada columna correspondiente de los
											// nombres a mostrar después del reemplazo
									rowElements.add(texto);
								}
							} else {

								if (row.getRowNum() == rowIndexStart) {
									// si la celda es de los nombres clave
									nombresReemplazo.add(texto); // para añadirlo a la lista de nombres a
									// reemplazar

								} else {
									rowElements.add(texto); // string vacío
								}
							}
						}
						if (rowElements.size() > 0) {
							rowList.add(rowElements);
						}
					}
				}
				// Se eliminan las posibles "palabras clave" vacías que pueda contener la tabla,
				// y con ello las columnas correspondientes ya que no interesarían
				for (int i = 0; i < nombresReemplazo.size(); i++) {
					if (nombresReemplazo.get(i).trim().length() == 0) {
						nombresReemplazo.remove(i);
						for (int j = 0; j < rowList.size(); j++) {
							rowList.get(j).remove(i);
						}
						i--;
					}
				}
				ds.setKeyNames(nombresReemplazo);
				ds.setRows(rowList);
				dSList.add(ds);
			}
			Controller.dataSources.set(dSList);
		}
		spreadSheet.close();
	}

	/**
	 * Lee el texto de una celda ofrecida en el método
	 * 
	 * @param c
	 * @return
	 */
	private String readCell(Cell c) {
		String texto = "";
		switch (c.getCellType()) {
		case STRING:
			texto = c.getStringCellValue();
			break;
		case BOOLEAN:
			texto = c.getStringCellValue();
			break;
		case FORMULA:
			texto = c.getCellFormula();
			break;
		case NUMERIC:
			texto = c.getNumericCellValue() + "";
		default:
			break;
		}
		return texto;
	}

	/**
	 * Convierte una lista de documentos en ficheros pdf
	 * 
	 * @param createdFiles Los documentos a convertir en pdf
	 * @param extension    La extensión del archivo que se quiere convertir a pdf
	 */
	private void convertFileListToPdfs(List<File> createdFiles, String extension) {
		ObservableList<File> files = FXCollections.observableArrayList();
		for (int i = 0; i < createdFiles.size(); i++) {
			String name = createdFiles.get(i).getPath();
			int lastIndex = name.lastIndexOf(extension);

			try {
				String beginName = name.substring(0, lastIndex);
				File f = new File(beginName + ".pdf");
				JodConverter.convert(createdFiles.get(i))// .as(DefaultDocumentFormatRegistry.DOC)
						.to(f)//
						.as(DefaultDocumentFormatRegistry.PDF)//
						.execute();
				files.add(f);
			} catch (OfficeException | StringIndexOutOfBoundsException e) {
			}
		}
		Controller.previsualizaciones.set(files);
	}
}
