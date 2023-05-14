package integracion.wordseedexporter.model;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import javax.xml.transform.TransformerException;
import javax.xml.xpath.XPathExpressionException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
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

public class DocumentManager {

	public DocumentManager() {
	}

	public void giveDocument(File f)
			throws XPathExpressionException, TransformerException, InvalidFormatException, IOException {

		if (f.getName().endsWith(".docx")) {
//			XWPFDocument doc = new XWPFDocument(OPCPackage.open(f));
//			document = doc;

			// ej.
			// Registros de la columna de "Nombres"
			List<String> nombres = Arrays.asList("Pepe", "Carlos"); // registros de una columna
			List<List<String>> listaClaves = Arrays.asList(nombres); // aquí irían más si hubiesen más columnas

			// Nombre de las columnas del excel - serán las palabras a reemplazar
			List<String> nombresColumnas = Arrays.asList("[Nombre]");
			replaceDocxStrings(listaClaves, nombresColumnas, f);
		}
	}

	// La lista de listas de string vendrá de la próxima interfaz a crear
	public void replaceDocxStrings(List<List<String>> replaceKeyList, List<String> keyList, File f)
			throws InvalidFormatException, IOException {

		// La interfaz IBody es la que implementa el método .getParagraphs() y las
		// clases que manejan el contenido de los párrafos
		// https://poi.apache.org/apidocs/dev/org/apache/poi/xwpf/usermodel/IBody.html
		// https://poi.apache.org/apidocs/dev/org/apache/poi/xwpf/usermodel/BodyType.html

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
						editParagraph(p, lChild.get(j), keyList.get(i));
					}
				}

				// Tablas
				if (tablas != null) {
					for (XWPFTable tbl : tablas) {
						for (XWPFTableRow row : tbl.getRows()) {
							for (XWPFTableCell cell : row.getTableCells()) {
								for (XWPFParagraph p : cell.getParagraphs()) {
									editParagraph(p, lChild.get(j), keyList.get(i));
								}
							}
						}
					}
				}

				// Comentarios
				if (comentarios != null) {
					for (XWPFComment comms : comentarios) {
						for (XWPFParagraph p : comms.getParagraphs()) {
							editParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// EndNotes
				if (endNotes != null) {
					for (XWPFEndnote endNote : endNotes) {
						for (XWPFParagraph p : endNote.getParagraphs()) {
							editParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Pies de páginas
				if (piesPag != null) {
					for (XWPFFooter footer : piesPag) {
						for (XWPFParagraph p : footer.getParagraphs()) {
							editParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Notas de pies de páginas
				if (notasPiesPag != null) {
					for (XWPFFootnote footNote : notasPiesPag) {
						for (XWPFParagraph p : footNote.getParagraphs()) {
							editParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}

				// Cabeceras
				if (cabeceras != null) {
					for (XWPFHeader header : cabeceras) {
						for (XWPFParagraph p : header.getParagraphs()) {
							editParagraph(p, lChild.get(j), keyList.get(i));
						}
					}
				}
				doc.write(new FileOutputStream("output_" + (j + 1) + ".docx"));
			}
		}

	}
	
	public void replacePptxStrings(File f) throws InvalidFormatException {
		XMLSlideShow slideShow = new XMLSlideShow(OPCPackage.open(f));
		
		// https://poi.apache.org/components/slideshow/quick-guide.html
		// https://poi.apache.org/apidocs/dev/org/apache/poi/xslf/usermodel/XSLFShapeContainer.html
		// https://poi.apache.org/apidocs/dev/org/apache/poi/xslf/usermodel/XSLFShape.html
		for (XSLFSlide slide : slideShow.getSlides()) {
			
			//slide.getNotes().getTextParagraphs().get(0).te
		}
	}

	public void editParagraph(XWPFParagraph p, String replaceableKey, String key) {
		List<XWPFRun> runs = p.getRuns();
		if (runs != null) {
			for (XWPFRun r : runs) {
				String text = r.getText(0);
//				if(text != null && text.equals("[Nombre]")) {
//					System.out.println(text);
//					System.out.println(replaceableKey);
//					System.out.println(key);
//				}
				if (text != null && text.contains(key)) {
					text = text.replace(key, replaceableKey);
					r.setText(text, 0);
				}
			}
		}
	}

}
