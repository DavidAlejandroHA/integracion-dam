package integracion.wordseedexporter.model;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
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
import org.odftoolkit.odfdom.doc.OdfPresentationDocument;
import org.odftoolkit.odfdom.doc.presentation.OdfSlide;
import org.odftoolkit.odfdom.dom.element.draw.DrawPageElement;
import org.w3c.dom.NodeList;

import integracion.wordseedexporter.controllers.Controller;
import javafx.collections.ObservableList;
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
	 * TODO: En caso de entregar un documento de la suite de OpenOffice o
	 * LibreOffice (odt, odp o ods), se transformará a su equivalente en las
	 * versiones de los documento de microsoft para poder editarlos utilizando las
	 * funciones actuales de la clase.
	 * 
	 * @param f El fichero (documento) a modificar
	 * @throws Exception
	 */
	public void giveDocument(File f) throws Exception {

		if (f != null) {
			// ej. // Registros de la columna de "Nombres"
			// List<String> nombres = Arrays.asList("Pepe", "Carlos"); // registros de una
			// columna
			// List<List<String>> listaClaves = Arrays.asList(nombres); // aquí irían más si
			// hubiesen más columnas
			// Nombre de las columnas del excel - serán las palabras a reemplazar
			// List<String> nombresColumnas = Arrays.asList("Sociedad");

			ObservableList<String> nombresColumnas = Controller.keyList.get();

			ObservableList<ObservableList<String>> columnas = Controller.columnList.get();

			if (f.getName().endsWith(".docx")) {
				replaceDocxStrings(nombresColumnas, columnas, f);

			} else if (f.getName().endsWith(".pptx")) {
				replacePptxStrings(nombresColumnas, columnas, f);

			} else if (f.getName().endsWith(".xlsx")) {
				replaceXlsxStrings(nombresColumnas, columnas, f);

			} else if (f.getName().endsWith(".odp")) {

				XPathFactoryUtil.setxPathFactory(new XPathFactoryImpl());
				if (Controller.keyList.get() != null && Controller.columnList.get() != null) {
					replaceOdpStrings(Controller.keyList.get(), Controller.columnList.get(), f);
				}

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

		// La interfaz IBody es la que implementa el método .getParagraphs() y las
		// clases que manejan el contenido de los párrafos en los docx

		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas

			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas

				XWPFDocument doc = new XWPFDocument(OPCPackage.open(f)); // para reiniciar el doc a su estado inicial
																			// para poder volver a reemplazar texto

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
						editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
					}
				}

				// Tablas
				if (tablas != null) {
					for (XWPFTable tbl : tablas) {
						for (XWPFTableRow row : tbl.getRows()) {
							for (XWPFTableCell cell : row.getTableCells()) {
								for (XWPFParagraph p : cell.getParagraphs()) {
									editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
								}
							}
						}
					}
				}

				// Comentarios
				if (comentarios != null) {
					for (XWPFComment comms : comentarios) {
						for (XWPFParagraph p : comms.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
						}
					}
				}

				// EndNotes
				if (endNotes != null) {
					for (XWPFEndnote endNote : endNotes) {
						for (XWPFParagraph p : endNote.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
						}
					}
				}

				// Pies de páginas
				if (piesPag != null) {
					for (XWPFFooter footer : piesPag) {
						for (XWPFParagraph p : footer.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
						}
					}
				}

				// Notas de pies de páginas
				if (notasPiesPag != null) {
					for (XWPFFootnote footNote : notasPiesPag) {
						for (XWPFParagraph p : footNote.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
						}
					}
				}

				// Cabeceras
				if (cabeceras != null) {
					for (XWPFHeader header : cabeceras) {
						for (XWPFParagraph p : header.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), columnKeyName.get(i));
						}
					}
				}
				doc.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + "_" + (i + 1) + ".docx"));
			}
		}
	}

	private void replacePptxStrings(ObservableList<String> columnKeyName,
			ObservableList<ObservableList<String>> columns, File f) throws InvalidFormatException, IOException {

		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas

			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				XMLSlideShow slideShow = new XMLSlideShow(OPCPackage.open(f));// para reiniciar el doc a su estado
																				// inicial para poder volver a
																				// reemplazar texto
				List<XSLFSlide> slides = slideShow.getSlides();
				if (slides != null) {
					for (XSLFSlide slide : slideShow.getSlides()) {
						// slide.getNotes().pa
						for (XSLFShape sh : slide.getShapes()) {
							if (sh instanceof XSLFTextShape) {

								List<XSLFTextParagraph> parrafos = ((XSLFTextShape) sh).getTextParagraphs();

								if (parrafos != null) {
									for (XSLFTextParagraph p : parrafos) {
										editPptxParagraph(p, lChild.get(j), columnKeyName.get(i));
									}
								}
							}
						}
					}
				}
				slideShow.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + "_" + (i + 1) + ".pptx"));

			}

		}
		// https://odftoolkit.org/simple/document/cookbook/Manipulate%20TextSearch.html
		// https://javadoc.io/static/org.odftoolkit/odfdom-java/0.11.0/org/odftoolkit/odfdom/doc/presentation/package-summary.html
		// https://javadoc.io/static/org.odftoolkit/odfdom-java/0.11.0/org/odftoolkit/odfdom/doc/package-summary.html
	}

	private void replaceXlsxStrings(ObservableList<String> columnKeyName,
			ObservableList<ObservableList<String>> columns, File f) throws InvalidFormatException, IOException {
		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				XSSFWorkbook spreadSheet = new XSSFWorkbook(OPCPackage.open(f));// para reiniciar el doc a su estado
																				// inicial para poder volver a
																				// reemplazar texto
				Iterator<Sheet> sheets = spreadSheet.sheetIterator();

				if (sheets != null) {

					// Para cada página del documento excel
					while (sheets.hasNext()) {
						Sheet sh = sheets.next(); // Se maneja cada hoja
						for (Row row : sh) {
							for (Cell cell : row) {
								editXlxsCells(cell, lChild.get(j), columnKeyName.get(i));
							}
						}
					}

				}
				spreadSheet.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + "_" + (i + 1) + ".pptx"));

			}
		}
	}

	private void replaceOdpStrings(ObservableList<String> columnKeyName, ObservableList<ObservableList<String>> columns,
			File f) throws Exception {

		for (int i = 0; i < columns.size(); i++) { // iterando en las columnas
			List<String> lChild = columns.get(i);
			for (int j = 0; j < lChild.size(); j++) { // iterando en las filas
				OdfPresentationDocument odpDocument = OdfPresentationDocument.loadDocument(f);

				// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
				// En este caso solo se escapea las comillas dobles porque es lo que se usa para
				// el contains"" del xpath
				String escapedKey1 = columnKeyName.get(i).replaceAll("\"", "\"\""); // Recibiendo el nombre de la
																					// columna + escapeando los
																					// caracteres

				String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";

				Iterator<OdfSlide> it = odpDocument.getSlides();
				while (it.hasNext()) {
					OdfSlide odfSlide = it.next();
					DrawPageElement slideElement = odfSlide.getOdfElement();
					NodeList slideNodeList = slideElement.getChildNodes();

					XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();
					NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(slideNodeList,
							XPathConstants.NODE);
					editStringNodeList(nodelist, lChild.get(j), columnKeyName.get(i));
				}
				odpDocument.save(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + "_" + (i + 1) + ".odp"));
			}
		}
	}

	private void editDocxParagraph(XWPFParagraph p, String newKey, String key) {
		String regexKey = stringModifyOptions(key);

		List<XWPFRun> runs = p.getRuns();
		if (runs != null) {
			for (XWPFRun r : runs) {
				String text = r.getText(0);
				if (text != null && text.contains(key)) {
					text = text.replaceAll(regexKey, newKey);
					r.setText(text, 0);
				}
			}
		}
	}

	private void editPptxParagraph(XSLFTextParagraph p, String newKey, String key) {
		String regexKey = stringModifyOptions(key);

		List<XSLFTextRun> runs = p.getTextRuns();
		if (runs != null) {
			for (XSLFTextRun r : runs) {
				String text = r.getRawText();
				if (text != null && text.contains(key)) {
					text = text.replaceAll(regexKey, newKey);
					r.setText(text);
				}
			}
		}
	}

	private void editXlxsCells(Cell c, String newKey, String key) {
		String regexKey = stringModifyOptions(key);

		// Solo modificar las celdas que contengan contenido de texto
		// Si contienen otras cosas que no sean texto, se ignoran ya que al modificar
		// por ejemplo celdas numéricas con texto, podría lanzarse una excepción
		if (c.getCellType().equals(CellType.STRING)) {
			String text = c.getStringCellValue();
			if (text != null && text.contains(key)) {
				text = text.replaceAll(regexKey, newKey);
				c.setCellValue(text);
			}
		}
	}

	private void editStringNodeList(NodeList nl, String newKey, String key) {
		if (nl != null) {
			// https://stackoverflow.com/questions/10664434/escaping-special-characters-in-java-regular-expressions
			//https://stackoverflow.com/questions/14134558/list-of-all-special-characters-that-need-to-be-escaped-in-a-regex
			for (int k = 0; k < nl.getLength(); k++) {
				if (nl.item(k).getTextContent().contains(key)) {
					String texto = nl.item(k).getTextContent();
					String controlRegex = stringModifyOptions(key);
					System.out.println("Texto clave - " + controlRegex);
					System.out.println("Texto nuevo - " + newKey);
					texto = texto.replaceAll(controlRegex, newKey);
					nl.item(k).setTextContent(texto);
				}
			}
		}
	}

	private String stringModifyOptions(String k) {
		// k = StringEscapeUtils.escapeJava(k); // es necesario el escape de caracteres
		// para que pueda trabajar bien con el replaceAll
		//System.out.println("antes - " + k);
		k = k.replaceAll("[\\<\\(\\[\\{\\\\\\^\\-\\=\\$\\!\\|\\]\\}\\)\\?\\*\\+\\.\\>]", "\\\\$0");
		if (Controller.replaceExactWord.get()) {
			// Si la BooleanProperty replaceExactWord está a true, se le añaden a la
			// replaceableKey (k) escapeada en Java un "limitador" de palabras para
			// seleccionar solo esas mismas palabras (\\b) y no otras que la puedan contener
			// pero que no sean la misma, ya que el .replaceAll() puede trabajar con
			// expresiones regulares (regex)
			k = "\\b" + k + "\\b";
		}
		return k;
	}
}
