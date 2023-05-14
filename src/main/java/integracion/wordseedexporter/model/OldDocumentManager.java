package integracion.wordseedexporter.model;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.xml.bind.JAXBException;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpressionException;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.utils.XPathFactoryUtil;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.JodConverter;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import integracion.wordseedexporter.WordSeedExporterApp;
import javafx.stage.FileChooser;
import net.sf.saxon.s9api.SaxonApiException;
import net.sf.saxon.xpath.XPathFactoryImpl;

public class OldDocumentManager {

	public OldDocumentManager() {

	}

	public void giveDocument(File f) throws XPathExpressionException, TransformerException, SaxonApiException {
		// https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/OdfTextDocument.html
		// https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/package-summary.html

		if (f.getName().endsWith(".docx")) {
			WordprocessingMLPackage wordMLPackage;

			try {
				wordMLPackage = WordprocessingMLPackage.load(f);
				MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

				// unmarshallFromTemplate requires string input
				String xml = XmlUtils.marshaltoString(documentPart.getContents(), true, true);

				xml = replaceStringInXML(xml, "'por'", "<PAPAPAPAs>&&apos;");

				// https://www.docx4java.org/blog/2019/01/opendope-and-xpath-2-03-0/
				// https://github.com/plutext/docx4j/blob/1bd5f8375525c2f3f22eef33077b593379391b77/src/samples/docx4j/org/docx4j/samples/ContentControlBindingExtensions.java
				Object obj = XmlUtils.unmarshalString(xml);

				// Inyectar el resultado en el docx
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
			} catch (SAXException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (ParserConfigurationException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}
		// https://docx4java.org/docx4j/Docx4j_GettingStarted.pdf
		// https://stackoverflow.com/questions/72405729/how-to-programatically-modify-libre-office-odt-document-in-java
		// https://angelozerr.wordpress.com/2012/12/06/how-to-convert-docxodt-to-pdfhtml-with-java/
		// https://stackoverflow.com/questions/61805246/docx4j-docx-to-pdf-conversion-docx-content-not-appearing-page-by-page-to-pdf
	}

	public String replaceStringInXML(String xml, String key1, String key2) throws SaxonApiException, SAXException,
			IOException, ParserConfigurationException, XPathExpressionException, TransformerException {

		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		factory.setNamespaceAware(true); // never forget this!
		DocumentBuilder docBuilder = factory.newDocumentBuilder();
		InputStream xmlStream = new ByteArrayInputStream(xml.getBytes());
		Document doc = docBuilder.parse(xmlStream);

		// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
		// En este caso solo se escapea las comillas dobles porque es lo que se usa para
		// el contains"" del xpath
		String escapedKey1 = key1.replaceAll("\"", "\"\"");

		// Añadir los carácteres de escape para la estructura xml (esto solo en archivos
		// xml de odf?)
		// key2 = StringEscapeUtils.escapeXml11(key2);

		String xpathExpression = "//*[text()[contains(.,\"" + escapedKey1 + "\")]]";
		// Para poder escapear los carácteres especiales
		XPathFactoryUtil.setxPathFactory(new XPathFactoryImpl());
		XPath xpath = XPathFactoryUtil.getXPathFactory().newXPath();

		NodeList nodelist = (NodeList) xpath.compile(xpathExpression).evaluate(doc, XPathConstants.NODE);
		// Actualizar los nodos encontrados
		// System.out.println(StringEscapeUtils.escapeXml11(key1));
		for (int i = 0; i < nodelist.getLength(); i++) {
			Node node = nodelist.item(i);
			String text = node.getTextContent();
			// Aqui ya no se utiliza el texto escapeado "de saxon" porque ya se está
			// editando texto a nivel de documento xml
			text = text.replaceAll(key1, key2);
			node.setTextContent(text);
		}
		// Escribir el contenido xml en forma de String
		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		DOMSource source = new DOMSource(doc);
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		StreamResult result = new StreamResult(output);
		transformer.transform(source, result);

		return new String(output.toByteArray());

//		Processor proc = new Processor(false);
//		Serializer serializer = proc.newSerializer();
//		XPathCompiler xpath = proc.newXPathCompiler();
//		DocumentBuilder builder = proc.newDocumentBuilder();
//
//		// Convertir la string a un único nodo padre (s9api)
//		XdmNode doc = builder.build(new StreamSource(new StringReader(xml)));
//
//		// Escapeando los carácteres de escapa para el XPath de Saxon (repetirlos)
//		List<String> escapeList = Arrays.asList("\"", "'", "<", ">", "&");
//		
//		for(int i = 0; i < escapeList.size(); i++) {
//			key1 = key1.replaceAll(escapeList.get(i), escapeList.get(i) + escapeList.get(i));
//			//key2 = key2.replaceAll(escapeList.get(i), escapeList.get(i) + escapeList.get(i));
//		}
//		
//		// Añadir los carácteres de escape para la estructura xml (esto solo en docx, pptx, exelx...)
//		key2 = StringEscapeUtils.escapeXml11(key2);
//
//		System.out.println(key1);
//		XPathCompiler co = proc.newXPathCompiler();
//		//co.evaluate("replace //*[text()[contains(.,\"" + key1 + "\")]], a", doc);
//		co.compile("replace //*[text()[contains(.,\"" + key1 + "\")]], a");
//		
//		//XPathSelector selector = xpath.compile("show //*[text()[contains(.,\"" + key1 + "\")]]").load();
//		
//		
//		for (XdmItem s : selector) {
//			//XdmNode ss = (XdmNode) s;
//			s.getUnderlyingValue().
//			System.out.println(s.getUnderlyingValue().getStringValue());
//		}
//		System.out.println("a");
//
//		serializer.setOutputProperty(Serializer.Property.OMIT_XML_DECLARATION, "no");
//		// serializer.setOutputProperty(Serializer.Property.INDENT, "yes");
//
//		return serializer.serializeNodeToString(doc);
	}

	public void convertDocToPDF(File f, String name) {
		FileChooser fileChooser = new FileChooser();
		//fileChooser.setInitialDirectory(new File("."));
		File saveFile = fileChooser.showSaveDialog(WordSeedExporterApp.primaryStage);
		try {
			File pdfFileOut = new File(saveFile.getPath() + File.separator + "output.pdf");
			JodConverter.convert(fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage))
					.as(DefaultDocumentFormatRegistry.DOCX).to(pdfFileOut).as(DefaultDocumentFormatRegistry.PDF)
					.execute();
		} catch (OfficeException e) {
			e.printStackTrace();
		} finally {
			System.out.println("b");
		}
	}
}
