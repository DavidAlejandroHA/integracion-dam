package integracion.wordseedexporter.model;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;

import javax.xml.bind.JAXBException;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.commons.io.IOUtils;
import org.apache.jena.rdfxml.xmlinput.impl.XMLHandler;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.incubator.search.TextNavigation;
import org.odftoolkit.odfdom.incubator.search.TextSelection;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

public class DocumentManager {
	
	public DocumentManager() {
		
	}
	
	public void giveDocument(File f) {
		//https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/OdfTextDocument.html
		//https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/package-summary.html // Tipos de documentos
		/*try {
			System.out.println("a");
			System.out.println(FileMagic.valueOf(f));
			//XMLSlideShow slideShow = new XMLSlideShow(new FileInputStream(f));
			System.out.println("b");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("sassdds");
		}*/
		
		if(f.getName().endsWith(".docx")) {
			WordprocessingMLPackage wordMLPackage;
			try {
				wordMLPackage = WordprocessingMLPackage
						.load(f);
				MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
				
				// unmarshallFromTemplate requires string input
				String xml = XmlUtils.marshaltoString(documentPart.getContents(), true, true);
				
				//System.out.println(xml);
				
				// TODO: hacer clase para parsear con el DOM y reemplazar solo el texto contenido en los elementos
				//
				//Document doc = convertStringToXMLDocument(xml);
				//xml = xml.replaceAll("por", "PAPAPAPA");
				xml = replaceWordInXML(xml, "por", "PAPAPAPA");
				// Do it...
				System.out.println(xml);
				Object obj = XmlUtils.unmarshalString(xml);
				// Inject result into docx
				//https://github.com/plutext/docx4j/blob/1f496eca1f70e07d8c112168857bee4c8e6b0514/docx4j-samples-docx4j/src/main/java/org/docx4j/samples/VariableReplace.java#L95
				//https://stackoverflow.com/questions/19325611/how-to-replace-text-in-content-control-after-xml-binding-using-docx4j
				documentPart.setJaxbElement((org.docx4j.wml.Document) obj);
				

				Save saver = new Save(wordMLPackage);
				saver.save(new FileOutputStream(new File("unmarshallFromTemplateExample.docx")));
				
			} catch (Docx4JException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (JAXBException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (ParserConfigurationException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (SAXException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		/*try {
			OdfTextDocument textdoc=(OdfTextDocument) OdfTextDocument.loadDocument(f);
			
			TextNavigation search1;

			search1= new TextNavigation("por",textdoc);

			while (search1.hasNext()) {

				TextSelection item1 = (TextSelection) search1.getCurrentItem();

				System.out.println(item1);

			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/
		//https://docx4java.org/docx4j/Docx4j_GettingStarted.pdf
		//https://stackoverflow.com/questions/72405729/how-to-programatically-modify-libre-office-odt-document-in-java
		//https://angelozerr.wordpress.com/2012/12/06/how-to-convert-docxodt-to-pdfhtml-with-java/
		//https://stackoverflow.com/questions/61805246/docx4j-docx-to-pdf-conversion-docx-content-not-appearing-page-by-page-to-pdf
	}
	
	public String replaceWordInXML(String xml, String key1, String key2) throws ParserConfigurationException, SAXException, IOException {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        SAXParser saxParser = factory.newSAXParser();
        InputStream targetStream = new ByteArrayInputStream(xml.getBytes());
        saxParser.parse(targetStream, new ReplaceWordHandler(key1, key2));
        System.out.println(IOUtils.toString(targetStream, StandardCharsets.UTF_8.name()));
		return IOUtils.toString(targetStream, StandardCharsets.UTF_8.name());
	}

}
