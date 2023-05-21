package integracion.wordseedexporter.model;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;

import org.apache.commons.io.FileUtils;
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

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.controllers.Controller;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import net.sf.saxon.xpath.XPathFactoryImpl;

/**
 * <p>
 * Esta clase se encarga de, dado un documento entregado, cambiar<br>
 * una palabra o una serie de palabras en el documento entregado.
 * </p>
 * <p>
 * Las palabras reemplazadas en el documento dependen además de la<br>
 * BooleanProperty de la clase Controller (Controller.replaceExactWord), <br>
 * ya que según su valor llegará a reemplazar solo las palabras <br>
 * exactas que el usuario haya especificado (true) o cualquier coincidencia<br>
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
	 * Este método gestiona el tipo de documento entregado (.docx, pptx, xlsx) y
	 * ejecuta el método apropiado para editar el documento.
	 * 
	 * @param f El fichero (documento) a modificar
	 * @throws Exception
	 */
	public void giveDocument(File f) throws Exception {

		if (f != null) {
			// Nombre de las columnas del excel - serán las palabras a reemplazar
			ObservableList<String> nombresColumnas = Controller.keyList.get();

			// Registros de la columna de Palabras (lista de columnas que contienen filas)
			ObservableList<ObservableList<String>> columnas = Controller.columnList.get();

			if ((columnas != null && nombresColumnas != null) && (columnas.size() > 0 && nombresColumnas.size() > 0)) {
				XPathFactoryUtil.setxPathFactory(new XPathFactoryImpl());

				if (f.getName().endsWith(".docx")) {
					replaceDocxStrings(nombresColumnas, columnas, f);

				} else if (f.getName().endsWith(".pptx")) {
					replacePptxStrings(nombresColumnas, columnas, f);

				} else if (f.getName().endsWith(".xlsx")) {
					replaceXlsxStrings(nombresColumnas, columnas, f);

				} else {

					if (f.getName().endsWith(".odt")) {
						replaceOdtStrings(nombresColumnas, columnas, f);

					} else if (f.getName().endsWith(".odp")) {
						replaceOdpStrings(nombresColumnas, columnas, f);

					} else if (f.getName().endsWith(".ods")) {
						replaceOdsStrings(nombresColumnas, columnas, f);

					} else if (f.getName().endsWith(".odg")) {
						replaceOdgStrings(nombresColumnas, columnas, f);

					}
				}
			}
		}
	}

	public void readData(File f) throws Exception {
		if (f.getName().endsWith(".ods")) {
			OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(f);
			List<OdfTable> tablas = odsDocument.getSpreadsheetTables();
			if (tablas != null) {
				String texto = null;
				ObservableList<String> filas = FXCollections.observableArrayList();
				ObservableList<String> nombresReemplazo = FXCollections.observableArrayList();

				ObservableList<ObservableList<String>> columnas = FXCollections
						.<ObservableList<String>>observableArrayList();

				for (OdfTable t : tablas) {
					for (int i = 0; i < t.getColumnCount(); i++) {
						filas = FXCollections.observableArrayList();
						for (int j = 0; j < t.getRowCount(); j++) {
							OdfTableCell cell = t.getCellByPosition(i, j);
							if (cell.getValueType() == null) {
								texto = "";
							} else {
								texto = cell.getStringValue();
							}
							if (j == 0) {
								nombresReemplazo.add(texto);
							} else {
								filas.add(texto);
							}
						}
						columnas.add(filas);
					}
				}
//				System.out.println("filas - " + filas);
//				System.out.println("columnas - " + columnas);
//				System.out.println("nombres - " + nombresReemplazo);
				Controller.columnList.set(columnas);
				Controller.keyList.set(nombresReemplazo);
			}
		}
	}

	// La lista de listas de string vendrá de la próxima interfaz a crear
	/**
	 * Edita el fichero (documento) entregado previamente por la función
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de creación de la fuente de datos o de
	 * la opción de importación.
	 * 
	 * @param columns
	 * @param columnKeyName
	 * @param f
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void replaceDocxStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		// La interfaz IBody es la que implementa el método .getParagraphs() y las
		// clases que manejan el contenido de los párrafos en los docx
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío
					XWPFDocument doc = new XWPFDocument(new FileInputStream(f)); // para reiniciar el doc a su estado
																					// inicial
																					// para poder volver a reemplazar
																					// texto
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
							cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
										cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
								cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
								cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
								cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
								cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
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
								cambios = editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
								if (cambios) {
									numCambios++;
								}
							}
						}
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".docx";
						doc.write(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					doc.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}

			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replacePptxStrings(ObservableList<String> columnKeyName,
			ObservableList<ObservableList<String>> columns, File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío
					XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(f));// para reiniciar el doc a su
																						// estado
																						// inicial para poder volver a
					boolean cambios = false;
					List<XSLFSlide> slides = slideShow.getSlides();
					if (slides != null) {
						for (XSLFSlide slide : slideShow.getSlides()) {
							for (XSLFShape sh : slide.getShapes()) {
								if (sh instanceof XSLFTextShape) {

									List<XSLFTextParagraph> parrafos = ((XSLFTextShape) sh).getTextParagraphs();

									if (parrafos != null) {
										for (XSLFTextParagraph p : parrafos) {
											cambios = editPptxParagraph(p, lChild.get(j), columnKeyName.get(i));
											if (cambios) {
												numCambios++;
											}
										}
									}
								}
							}
						}
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".pptx";
						slideShow.write(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}

					slideShow.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}

			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replaceXlsxStrings(ObservableList<String> columnKeyName,
			ObservableList<ObservableList<String>> columns, File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0)
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {
					XSSFWorkbook spreadSheet = new XSSFWorkbook(new FileInputStream(f));

					boolean cambios = false;
					Iterator<Sheet> sheets = spreadSheet.sheetIterator();

					if (sheets != null) {

						// Para cada página del documento excel
						while (sheets.hasNext()) {
							Sheet sh = sheets.next(); // Se maneja cada hoja
							for (Row row : sh) {
								for (Cell cell : row) {
									cambios = editXlxsCells(cell, lChild.get(j), columnKeyName.get(i));
									if (cambios) {
										numCambios++;
									}
								}
							}
						}

					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".pptx";
						spreadSheet.write(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					spreadSheet.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replaceOdtStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío

					OdfTextDocument odtDocument = OdfTextDocument.loadDocument(f);
					boolean cambios = false;
					// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
					// En este caso solo se escapea las comillas dobles porque es lo que se usa para
					// el contains"" del xpath
					String escapedKey1 = columnKeyName.get(i).replaceAll("\"", "\"\""); // Recibiendo el nombre de la
																						// columna + escapeando las
																						// comillas

					String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

					OfficeTextElement odtTextEl = odtDocument.getContentRoot();
					NodeList odtNodes = odtTextEl.getChildNodes();

					XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
					NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
							XPathConstants.NODE);
					cambios = editStringNodeList(nodelist, lChild.get(j), columnKeyName.get(i));
					if (cambios) {
						numCambios++;
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".odt";
						odtDocument.save(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					odtDocument.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replaceOdpStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío
					OdfPresentationDocument odpDocument = OdfPresentationDocument.loadDocument(f);
					boolean cambios = false;
					// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
					// En este caso solo se escapea las comillas dobles porque es lo que se usa para
					// el contains"" del xpath
					String escapedKey1 = columnKeyName.get(i).replaceAll("\"", "\"\""); // Recibiendo el nombre de la
																						// columna + escapeando las
																						// barras
					String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

					Iterator<OdfSlide> it = odpDocument.getSlides();
					while (it.hasNext()) {
						OdfSlide odfSlide = it.next();
						DrawPageElement slideElement = odfSlide.getOdfElement();
						NodeList slideNodeList = slideElement.getChildNodes();

						XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
						NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(slideNodeList,
								XPathConstants.NODE);
						cambios = editStringNodeList(nodelist, lChild.get(j), columnKeyName.get(i));
						if (cambios) {
							numCambios++;
						}
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".odp";
						odpDocument.save(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					odpDocument.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replaceOdsStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío
					OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(f);
					boolean cambios = false;
					// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
					// En este caso solo se escapea las comillas dobles porque es lo que se usa para
					// el contains"" del xpath
					String escapedKey1 = columnKeyName.get(i).replaceAll("\"", "\"\""); // Recibiendo el nombre de la
																						// columna + escapeando las
																						// comillas

					String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";
					OfficeSpreadsheetElement odtSShEl = odsDocument.getContentRoot();
					NodeList odtNodes = odtSShEl.getChildNodes();

					XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
					// ((XPathEvaluator)xpath).getStaticContext().setDefaultElementNamespace(NamespaceUri.getUriForConventionalPrefix("http://www.w3.org/1999/xhtml"));
					NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
							XPathConstants.NODE);
					cambios = editStringNodeList(nodelist, lChild.get(j), columnKeyName.get(i));
					if (cambios) {
						numCambios++;
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".ods";
						odsDocument.save(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					odsDocument.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

	private void replaceOdgStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws Exception {
		int iEffective = 0;
		int jEffective = 0;
		int numCambios = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			iEffective = 0;
			numCambios = 0;
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				if ((columnKeyName.get(i) != null && columnKeyName.get(i).trim().length() > 0) // si ninguno de los 2
						&& (lChild.get(j) != null && lChild.get(j).trim().length() > 0)) {// registros está vacío
					OdfGraphicsDocument odgDocument = OdfGraphicsDocument.loadDocument(f);
					boolean cambios = false;
					// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
					// En este caso solo se escapea las comillas dobles porque es lo que se usa para
					// el contains"" del xpath
					String escapedKey1 = columnKeyName.get(i).replaceAll("\"", "\"\""); // Recibiendo el nombre de la
																						// columna + escapeando las
																						// comillas

					String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

					OfficeDrawingElement odgDrawEl = odgDocument.getContentRoot();
					NodeList odtNodes = odgDrawEl.getChildNodes();

					XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
					NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(odtNodes,
							XPathConstants.NODE);
					cambios = editStringNodeList(nodelist, lChild.get(j), columnKeyName.get(i));
					if (cambios) {
						numCambios++;
					}
					if (numCambios > 0) {
						String fileName = "output_" + (jEffective + 1) + "_" + (iEffective + 1) + ".odg";
						odgDocument.save(
								new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
						createdFiles.add(fileName);
						iEffective++;
					}
					odgDocument.close();
				}
			}
			boolean checkColumn = false;
			for (String s : lChild) {
				if (s != null && s.trim().length() > 0) {
					checkColumn = true;
				}
			}
			if (checkColumn) {
				jEffective++;
			}
			if (numCambios == 0) {
				notFound.add(columnKeyName.get(i));
			}
		}

		// Eliminar los ficheros que pueden haber quedado de acciones pasadas, sin
		// incluir los pdf
		deleteOldFiles(createdFiles);

		if (notFound.size() > 0) { // si hay palabras-clave que no se han encontrado en el documento
			notFoundWords(notFound);
		} else { // si no las hay es que se han generado todos los documentos correspondientes
			allWordsSuccess();
		}
	}

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
					text = text.replaceAll(regexKey, newKey);
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
					text = text.replaceAll(regexKey, newKey);
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
				text = text.replaceAll(regexKey, newKey);
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

	private boolean editStringNodeList(NodeList nl, String newKey, String key) {
		boolean cambios = false;
		int numCambios = 0;
		if (nl != null) {
			for (int k = 0; k < nl.getLength(); k++) {
				if (nl.item(k).getTextContent().contains(key)) {
					String texto = nl.item(k).getTextContent();
					String textAux = texto;
					String controlRegex = stringModifyOptions(key);
					texto = texto.replaceAll(controlRegex, newKey);
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

	private String stringModifyOptions(String k) {
		k = Pattern.quote(k); // con este método se obtiene un patrón "literal" del string, deforma que los
								// caracteres especiales que puedatener el string quedan "asegurados" en lo que
								// se refiere a lasintaxis de las expresiones regulares que utiliza el
								// método.replaceAll(), evitando así posible errores de sintaxis en
								// estas expresiones
								// Ojo: no es lo mismo que Matcher.quoteReplacement(k)
		// System.out.println(k);
		if (Controller.replaceExactWord.get()) {
			// Si la BooleanProperty replaceExactWord está a true, se le añaden a k un
			// "limitador" de palabras para seleccionar solo esas mismas palabras (.* + k +
			// *.) y no otras que la puedan contener pero que no sean la misma, ya que el
			// .replaceAll() puede trabajar con expresiones regulares (regex)
			k = "\\b" + k + "\\b";
		}
		return k;
	}

	private void notFoundWords(List<String> ls) {
		Alert alerta = new Alert(AlertType.WARNING);
		alerta.setTitle("No se han encontrado algunas palabras");
		String textoPalabras = "";
		for (int i = 0; i < ls.size(); i++) {
			textoPalabras += ((i < ls.size() - 1) ? "" : "y ") + "\"" + ls.get(i) + "\""
					+ ((i < ls.size() - 1) ? ", " : " ");
		}
		alerta.setHeaderText("Las palabras " + textoPalabras + "no se han encontrado en el fichero importado.");
		alerta.setContentText(
				"Varias palabras especificadas en la fuente de datos no han sido encontradas, por lo que no "
						+ "se han aplicado los cambios correspondientes respecto a estas palabras.");
		alerta.initOwner(WordSeedExporterApp.primaryStage);
		alerta.showAndWait();
	}

	private void allWordsSuccess() {
		Alert alerta = new Alert(AlertType.INFORMATION);
		alerta.setTitle("Operación exitosa");
		alerta.setHeaderText(
				"Se han reemplazado todas las palabras de la fuente de datos y aplicado los cambios correspondientes.");
		alerta.initOwner(WordSeedExporterApp.primaryStage);
		alerta.showAndWait();
	}

	private void deleteOldFiles(List<String> ls) {
		List<File> fileList = new ArrayList<File>(
				FileUtils.listFiles(new File(Controller.TEMPDOCSFOLDER.getPath()), null, false));
		for (File file : fileList) {
			if (!ls.contains(file.getName()) && !file.getName().endsWith(".pdf")) {
				FileUtils.deleteQuietly(file);
			}
		}
	}
}
