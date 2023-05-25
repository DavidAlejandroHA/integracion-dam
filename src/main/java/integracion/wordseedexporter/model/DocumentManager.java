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
 * Esta clase se encarga de, dado un documento entregado, generar distintos<br>
 * documentos, cada uno con los cambios de palabras correspondientes<br>
 * respecto a la fuente de datos previamente importada en la aplicación<br>
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
	 * Este método gestiona el tipo de documento entregado (.docx, pptx, xlsx, odt, odp, ods o odg) y
	 * ejecuta el método apropiado para editar el documento.
	 * 
	 * @param f El fichero (documento) a modificar
	 * @throws Exception
	 */
	public void giveDocument(File f) throws Exception {

		if (f != null) {
			// Nombre de las columnas del excel - serán las palabras a reemplazar
			ObservableList<String> palabrasClave = Controller.keyList.get();

			// Registros de las filas de palabras para nuevos valores
			ObservableList<ObservableList<String>> filas = Controller.rowList.get();

			if ((filas != null && palabrasClave != null) && (filas.size() > 0 && palabrasClave.size() > 0)) {
				XPathFactoryUtil.setxPathFactory(new XPathFactoryImpl());

				if (f.getName().endsWith(".docx")) {
					replaceDocxStrings(palabrasClave, filas, f);

				} else if (f.getName().endsWith(".pptx")) {
					replacePptxStrings(palabrasClave, filas, f);

				} else if (f.getName().endsWith(".xlsx")) {
					replaceXlsxStrings(palabrasClave, filas, f);

				} else {

					if (f.getName().endsWith(".odt")) {
						replaceOdtStrings(palabrasClave, filas, f);

					} else if (f.getName().endsWith(".odp")) {
						replaceOdpStrings(palabrasClave, filas, f);

					} else if (f.getName().endsWith(".ods")) {
						replaceOdsStrings(palabrasClave, filas, f);

					} else if (f.getName().endsWith(".odg")) {
						replaceOdgStrings(palabrasClave, filas, f);

					}
				}
			}
		}

	}

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
	 * Edita el fichero (documento) entregado previamente por la función
	 * {@link #giveDocument(File)} junto a la lista de palabras clave y la lista de
	 * palabras a reemplazar que contendrán en sus índices correspondientes los
	 * Strings almacenados a través del menú de creación de la fuente de datos o de
	 * la opción de importación.
	 * 
	 * @param rows
	 * @param columnKeyName
	 * @param f
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void replaceDocxStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		// La interfaz IBody es la que implementa el método .getParagraphs() y las
		// clases que manejan el contenido de los párrafos en los docx
		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);
			XWPFDocument doc = new XWPFDocument(new FileInputStream(f)); // para reiniciar el doc a su estado inicial
																			// para poder volver a reemplazar texto

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

					if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se vuelve
																						// a agregar
						notFound.add(columnKeyName.get(j));
					}
				}
			}

			if (numCambiosRow > 0) {
				iEffective++;
				String fileName = "output_" + (iEffective) + ".docx";
				doc.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);
			}
			doc.close();
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

	private void replacePptxStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);
			XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(f)); // para reiniciar el doc a su estado
																				// inicial
																				// para poder volver a reemplazar texto

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

					if (numCambios == 0 && !notFound.contains(columnKeyName.get(j))) { // si ya lo contiene no se vuelve
																						// a agregar
						notFound.add(columnKeyName.get(j));
					}
				}
			}

			if (numCambiosRow > 0) {
				System.out.println("cambiosrow");
				iEffective++;
				String fileName = "output_" + (iEffective) + ".pptx";
				slideShow.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);
			}
			slideShow.close();
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

	private void replaceXlsxStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws InvalidFormatException, IOException {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);

			XSSFWorkbook spreadSheet = new XSSFWorkbook(new FileInputStream(f)); // para reiniciar el doc a su estado
																					// inicial
																					// para poder volver a reemplazar
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
				String fileName = "output_" + (iEffective) + ".xlsx";
				spreadSheet
						.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);
			}
			spreadSheet.close();
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

	private void replaceOdtStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws Exception {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);

			OdfTextDocument odtDocument = OdfTextDocument.loadDocument(f); // para reiniciar el doc a su estado
																			// inicial
																			// para poder volver a reemplazar
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
				String fileName = "output_" + (iEffective) + ".odt";
				odtDocument.save(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);

			}
			odtDocument.close();
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

	private void replaceOdpStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws Exception {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);

			OdfPresentationDocument odpDocument = OdfPresentationDocument.loadDocument(f); // para reiniciar el doc a su
																							// estado
			// inicial
			// para poder volver a reemplazar
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
				String fileName = "output_" + (iEffective) + ".odp";
				odpDocument.save(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);

			}
			odpDocument.close();
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

	private void replaceOdsStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws Exception {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);

			OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(f); // para reiniciar el doc a su
																							// estado
			// inicial
			// para poder volver a reemplazar
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
				String fileName = "output_" + (iEffective) + ".ods";
				odsDocument.save(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);

			}
			odsDocument.close();
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

	private void replaceOdgStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> rows,
			File f) throws Exception {
		int iEffective = 0;
		int numCambios = 0;
		int numCambiosRow = 0;
		List<String> notFound = new ArrayList<>();
		List<String> createdFiles = new ArrayList<>();

		for (int i = 0; i < rows.size(); i++) { // iterando en las filas
			numCambiosRow = 0;

			List<String> cellString = rows.get(i);

			OdfGraphicsDocument odgDocument = OdfGraphicsDocument.loadDocument(f); // para reiniciar el doc a su estado
			// inicial
			// para poder volver a reemplazar
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
				String fileName = "output_" + (iEffective) + ".odg";
				odgDocument.save(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + fileName));
				createdFiles.add(fileName);

			}
			odgDocument.close();
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
				"Varias palabras especificadas en la fuente de datos no han sido encontradas, por lo que no\n"
						+ "se han aplicado los cambios correspondientes respecto a estas palabras.");
		alerta.initOwner(WordSeedExporterApp.primaryStage);
		alerta.showAndWait();
	}

	private void allWordsSuccess() {
		Alert alerta = new Alert(AlertType.INFORMATION);
		alerta.setTitle("Operación exitosa");
		alerta.setHeaderText("Se han reemplazado todas las palabras de la fuente de datos\n"
				+ "y aplicado los cambios correspondientes.");
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

	private void readOds(File f) throws Exception {
		OdfSpreadsheetDocument odsDocument = OdfSpreadsheetDocument.loadDocument(f);
		List<OdfTable> tablas = odsDocument.getSpreadsheetTables();
		if (tablas != null) {

			ObservableList<String> celdas = FXCollections.observableArrayList();
			ObservableList<String> nombresReemplazo = FXCollections.observableArrayList();

			ObservableList<ObservableList<String>> filas = FXCollections.<ObservableList<String>>observableArrayList();

			// leer el tamaño de la tabla
			int width = 0;
			int height = 0;
			int rowIndexStart = 0;
			int columnIndexStart = 0;
			boolean lock = false;

			for (OdfTable t : tablas) {

				// lecutra de las dimensiones de la tabla y su punto de partida
				for (int i = 0; i < t.getRowCount(); i++) {
					boolean contains = false;

					celdas = FXCollections.observableArrayList(); // reset de filas
					for (int j = 0; j < t.getColumnCount(); j++) {
						OdfTableCell cell = t.getCellByPosition(j, i);

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
					throw new Exception();
				}

				String texto = null;
				for (int i = rowIndexStart; i < rowIndexStart + height; i++) {
					celdas = FXCollections.observableArrayList(); // reset de filas
					for (int j = columnIndexStart; j < columnIndexStart + width; j++) {
						OdfTableCell cell = t.getCellByPosition(j, i);
						if (cell == null) { // si hay datos el texto son los datos de la casilla
							texto = ""; // si no se pone a string vacío
						} else {
							texto = cell.getStringValue();
						}
						if (i == rowIndexStart) { // si es el primer registro se usa como una palabra a reemplazar en el
							// documento
							// if (texto != null && texto.trim().length() > 0) {
							nombresReemplazo.add(texto);
							// }
						} else { // si no como una de las claves a usar para el reemplazo de palabras
							celdas.add(texto);
						}
					}
					if (celdas.size() > 0) { // celdas = 1 fila
						filas.add(celdas); // una vez procesadas todas las filas se añaden a la lista de filas
					}
				}
			}

			// eliminar las columnas que no tienen nombres clave
			for (int i = 0; i < nombresReemplazo.size(); i++) {
				// System.out.println(nombresReemplazo.get(i));
				if (nombresReemplazo.get(i).trim().length() == 0) {
					nombresReemplazo.remove(i);
					for (int j = 0; j < filas.size(); j++) {
						filas.get(j).remove(i);
					}
					i--;
				}
			}

			// En esta última parte, en caso de leer varias hojas, es posible que las tablas
			// puedan tener distintos tamaños, por lo que una vez juntas todas las tablas y
			// las filas, se le añaden strings vacíos que no influirán en el reemplazo de
			// palabras a cada fila que sea necesaria para que todas tengan el mismo tamaño
			// y no se produzca posteriormente un IndexOutOfBoundException
			int topRowLength = 0;
			for (int i = 0; i < filas.size(); i++) {
				for (int j = 0; j < filas.get(i).size(); j++) {
					if (filas.get(i).size() > topRowLength) {
						topRowLength = filas.get(i).size();
					}
				}

				int missingElemNum = topRowLength - filas.get(i).size();
				if (missingElemNum != 0) {
					for (int k = 0; k < missingElemNum; k++) {
						filas.get(i).add("");
					}
				}
			}

			Controller.rowList.set(filas);
			Controller.keyList.set(nombresReemplazo);
			// System.out.println(filas);
			// System.out.println(nombresReemplazo);
		}
	}

	private void readXlsx(File f) throws Exception {
		XSSFWorkbook spreadSheet = new XSSFWorkbook(new FileInputStream(f));

		Iterator<Sheet> sheets = spreadSheet.sheetIterator();

		if (sheets != null) {

			ObservableList<String> rowElements = FXCollections.observableArrayList();
			ObservableList<String> nombresReemplazo = FXCollections.observableArrayList();

			ObservableList<ObservableList<String>> filas = FXCollections.<ObservableList<String>>observableArrayList();

			ObservableList<ObservableList<String>> rowList = FXCollections.observableArrayList();
			// Para cada página del documento excel
			while (sheets.hasNext()) {
				Sheet sh = sheets.next(); // Se maneja cada hoja
				// https://poi.apache.org/components/spreadsheet/quick-guide.html#TextExtraction

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
					throw new Exception();
				}

				for (int i = rowIndexStart; i < rowIndexStart + height; i++) { // manejando cada fila
					rowElements = FXCollections.observableArrayList();
					Row row = sh.getRow(i);

					if (row != null) {
						for (int cn = columnIndexStart; cn < columnIndexStart + width; cn++) { // manejandp
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
			}

			filas.setAll(rowList);

			// Se eliminan las posibles "palabras clave" vacías que pueda contener la tabla,
			// y con ello las columnas correspondientes ya que no interesarían
			for (int i = 0; i < nombresReemplazo.size(); i++) {
				// System.out.println(nombresReemplazo.get(i));
				if (nombresReemplazo.get(i).trim().length() == 0) {
					nombresReemplazo.remove(i);
					for (int j = 0; j < filas.size(); j++) {
						filas.get(j).remove(i);
					}
					i--;
				}
			}

			// En esta última parte, en caso de leer varias hojas, es posible que las tablas
			// puedan tener distintos tamaños, por lo que una vez juntas todas las tablas y
			// las filas, se le añaden strings vacíos que no influirán en el reemplazo de
			// palabras a cada fila que sea necesaria para que todas tengan el mismo tamaño
			// y no se produzca posteriormente un IndexOutOfBoundException
			int topRowLength = 0;
			for (int i = 0; i < filas.size(); i++) {
				for (int j = 0; j < filas.get(i).size(); j++) {
					if (filas.get(i).size() > topRowLength) {
						topRowLength = filas.get(i).size();
					}
				}

				int missingElemNum = topRowLength - filas.get(i).size();
				if (missingElemNum != 0) {
					for (int k = 0; k < missingElemNum; k++) {
						filas.get(i).add("");
					}
				}
				System.out.println(missingElemNum);
			}

			System.out.println(filas);
			System.out.println(nombresReemplazo);
			Controller.rowList.set(filas);
			Controller.keyList.set(nombresReemplazo);
		}
	}

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

//	private <T> ObservableList<ObservableList<T>> transpose(ObservableList<ObservableList<T>> table) {
//		ObservableList<ObservableList<T>> result = FXCollections.observableArrayList();
//		final int N = table.get(0).size();
//		for (int i = 0; i < N; i++) {
//			ObservableList<T> col = FXCollections.observableArrayList();
//			for (ObservableList<T> row : table) {
//				col.add(row.get(i));
//			}
//			result.add(col);
//		}
//		return result;
//	}
}
