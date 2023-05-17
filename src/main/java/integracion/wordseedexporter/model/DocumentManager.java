package integracion.wordseedexporter.model;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.text.StringEscapeUtils;
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

import integracion.wordseedexporter.controllers.Controller;
import javafx.collections.ObservableList;

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
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void giveDocument(File f) throws InvalidFormatException, IOException {

		if (f != null) {
			// ej. // Registros de la columna de "Nombres"
			// List<String> nombres = Arrays.asList("Pepe", "Carlos"); // registros de una
			// columna
			// List<List<String>> listaClaves = Arrays.asList(nombres); // aquí irían más si
			// hubiesen más columnas

			ObservableList<String> listaClaves = Controller.keyList.get();

			// Nombre de las columnas del excel - serán las palabras a reemplazar
			// List<String> nombresColumnas = Arrays.asList("Sociedad");

			ObservableList<ObservableList<String>> nombresColumnas = Controller.columnList.get();

			if (f.getName().endsWith(".docx")) {
				replaceDocxStrings(listaClaves, nombresColumnas, f);

			} else if (f.getName().endsWith(".pptx")) {
				replacePptxStrings(listaClaves, nombresColumnas, f);

			} else if (f.getName().endsWith(".xlsx")) {
				replaceXlsxStrings(listaClaves, nombresColumnas, f);
			} else if (f.getName().endsWith(".odp")) {

//				try {
//					//OdfTextDocument document = OdfTextDocument.loadDocument(f);
//					OdfPresentationDocument document = OdfPresentationDocument.loadDocument(f);
//			        // 5) Convert ODFDOM OdfTextDocument 2 PDF with iText
//			        OutputStream salida = new FileOutputStream(new File("test2.pdf"));
//			        PdfOptions options = PdfOptions.create();
//			        PdfConverter.getInstance().convert(document, salida, options);
//				} catch (Exception e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
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
	 * @param replaceKeyList
	 * @param keyList
	 * @param f
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public void replaceDocxStrings(ObservableList<String> keyList,
			ObservableList<ObservableList<String>> replaceKeyList, File f) throws InvalidFormatException, IOException {

		// La interfaz IBody es la que implementa el método .getParagraphs() y las
		// clases que manejan el contenido de los párrafos en los docx

		for (int i = 0; i < replaceKeyList.size(); i++) { // iterando en las columnas

			List<String> lChild = replaceKeyList.get(i);
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
						editDocxParagraph(p, lChild.get(j), keyList.get(i));
					}
				}

				// Tablas
				if (tablas != null) {
					for (XWPFTable tbl : tablas) {
						for (XWPFTableRow row : tbl.getRows()) {
							for (XWPFTableCell cell : row.getTableCells()) {
								for (XWPFParagraph p : cell.getParagraphs()) {
									editDocxParagraph(p, lChild.get(j), keyList.get(i));
								}
							}
						}
					}
				}

				// Comentarios
				if (comentarios != null) {
					for (XWPFComment comms : comentarios) {
						for (XWPFParagraph p : comms.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// EndNotes
				if (endNotes != null) {
					for (XWPFEndnote endNote : endNotes) {
						for (XWPFParagraph p : endNote.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Pies de páginas
				if (piesPag != null) {
					for (XWPFFooter footer : piesPag) {
						for (XWPFParagraph p : footer.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Notas de pies de páginas
				if (notasPiesPag != null) {
					for (XWPFFootnote footNote : notasPiesPag) {
						for (XWPFParagraph p : footNote.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Cabeceras
				if (cabeceras != null) {
					for (XWPFHeader header : cabeceras) {
						for (XWPFParagraph p : header.getParagraphs()) {
							editDocxParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}
				doc.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + (i + 1) + ".docx"));
			}
		}
	}

	public void replacePptxStrings(ObservableList<String> keyList,
			ObservableList<ObservableList<String>> replaceKeyList, File f) throws InvalidFormatException, IOException {

		for (int i = 0; i < replaceKeyList.size(); i++) { // iterando en las columnas

			List<String> lChild = replaceKeyList.get(i);
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
										editPptxParagraph(p, lChild.get(j), keyList.get(i));
									}
								}
							}
						}
					}
				}
				slideShow.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + (i + 1) + ".pptx"));

			}

		}
		// https://odftoolkit.org/simple/document/cookbook/Manipulate%20TextSearch.html
		// https://javadoc.io/static/org.odftoolkit/odfdom-java/0.11.0/org/odftoolkit/odfdom/doc/presentation/package-summary.html
		// https://javadoc.io/static/org.odftoolkit/odfdom-java/0.11.0/org/odftoolkit/odfdom/doc/package-summary.html
	}

	public void replaceXlsxStrings(ObservableList<String> keyList,
			ObservableList<ObservableList<String>> replaceKeyList, File f) throws InvalidFormatException, IOException {
		for (int i = 0; i < replaceKeyList.size(); i++) { // iterando en las columnas
			List<String> lChild = replaceKeyList.get(i);
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
								editXlxsCells(cell, lChild.get(j), keyList.get(i));
							}
						}
					}

				}
				spreadSheet.write(new FileOutputStream(Controller.TEMPDOCSFOLDER.getPath() + File.separator + "output_"
						+ (j + 1) + (i + 1) + ".pptx"));

			}
		}
	}

	private void editDocxParagraph(XWPFParagraph p, String replaceableKey, String key) {
		replaceableKey = stringModifyOptions(replaceableKey);

		List<XWPFRun> runs = p.getRuns();
		if (runs != null) {
			for (XWPFRun r : runs) {
				String text = r.getText(0);
				if (text != null && text.contains(key)) {
					text = text.replaceAll(key, replaceableKey);
					r.setText(text, 0);
				}
			}
		}
	}

	private void editPptxParagraph(XSLFTextParagraph p, String replaceableKey, String key) {
		replaceableKey = stringModifyOptions(replaceableKey);

		List<XSLFTextRun> runs = p.getTextRuns();
		if (runs != null) {
			for (XSLFTextRun r : runs) {
				String text = r.getRawText();
				if (text != null && text.contains(key)) {
					text = text.replaceAll(key, replaceableKey);
					r.setText(text);
				}
			}
		}
	}

	private void editXlxsCells(Cell c, String replaceableKey, String key) {
		replaceableKey = stringModifyOptions(replaceableKey);

		// Solo modificar las celdas que contengan contenido de texto
		// Si contienen otras cosas que no sean texto, se ignoran ya que al modificar
		// por ejemplo celdas numéricas con texto, podría lanzarse una excepción
		if (c.getCellType().equals(CellType.STRING)) {
			String text = c.getStringCellValue();
			if (text != null && text.contains(key)) {
				text = text.replaceAll(key, replaceableKey);
				c.setCellValue(text);
			}
		}
	}

	private String stringModifyOptions(String k) {
		k = StringEscapeUtils.escapeJava(k); // es necesario el escape de caracteres para que
		// pueda trabajar bien con el replaceAll
		if (Controller.replaceExactWord.get()) {
			// Si la BooleanProperty replaceExactWord está a true, se le añaden a la
			// replaceableKey
			// (k) escapeada en Java un "limitador" de palabras para seleccionar solo esas
			// mismas palabras (\\b) y no otras que la puedan contener pero que no sean la
			// misma, ya que el .replaceAll() puede trabajar con expresiones regulares
			// (regex)
			k = "\\b" + k + "\\b";
		}
		return k;
	}
}
