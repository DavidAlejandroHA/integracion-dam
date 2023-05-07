package integracion.informgenerator.controllers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.ocr.TesseractOCRConfig;
import org.apache.tika.parser.pdf.PDFParserConfig;
import org.apache.tika.sax.ContentHandlerDecorator;
import org.apache.tika.sax.ToXMLContentHandler;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;

import com.jfoenix.controls.JFXButton;

import integracion.informgenerator.InformGeneratorApp;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ButtonType;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class VBoxDrawerController implements Initializable {

	// view

	@FXML
	private HBox drawerView;

	@FXML
	private JFXButton importarDocumentoButton;

	@FXML
	private JFXButton am;

	// model

	WordToHtmlConverter wordToHtmlConverter;
	HWPFDocumentCore wordDocument;

	@Override
	public void initialize(URL location, ResourceBundle resources) {

	}

	public VBoxDrawerController() {
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/DrawerMenuView.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public HBox getView() {
		return drawerView;
	}

	@FXML
	void dda(ActionEvent event) {
		System.out.println("asasasas2");
	}

	@FXML
	void sas(MouseEvent event) {
		System.out.println("asasasas333313321");
	}

	@FXML
	void importarDocumento(ActionEvent event) {
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));
		// https://jenkov.com/tutorials/javafx/filechooser.html
		// https://blog.csdn.net/u011781521/article/details/116260048
		/*
		 * try { System.out.println(FileMagic.valueOf(fileChooser.showOpenDialog(
		 * InformGeneratorApp.primaryStage))); } catch (IOException e1) { // TODO
		 * Auto-generated catch block e1.printStackTrace(); }
		 */
		try {

			AutoDetectParser parser = new AutoDetectParser();
			Metadata metadata = new Metadata();
			ContentHandler handler = new ToXMLContentHandler();

			/*ContentHandlerDecorator h2Handler = new ContentHandlerDecorator(handler) {
				private final List<String> elementsToInclude = List.of("head");

				@Override
				public void startElement(String uri, String local, String name, Attributes atts) throws SAXException {
					if (elementsToInclude.contains(name)) {
						AttributesImpl s = new AttributesImpl();
						s.addAttribute(uri, local, name, local, name);
						super.startElement(uri, local, name, s);
					}
				}

				@Override
				public void endElement(String uri, String local, String name) throws SAXException {
					if (elementsToInclude.contains(name)) {
						super.endElement(uri, local, name);
					}
				}
			};*/

			TesseractOCRConfig config = new TesseractOCRConfig();
			PDFParserConfig pdfConfig = new PDFParserConfig();
			pdfConfig.setExtractInlineImages(true);

			// To parse images in files those lines are needed
			ParseContext parseContext = new ParseContext();
			parseContext.set(TesseractOCRConfig.class, config);
			parseContext.set(PDFParserConfig.class, pdfConfig);
			parseContext.set(Parser.class, parser); // need to add this to make sure
													// recursive parsing happens!

			InputStream stream = new FileInputStream(fileChooser.showOpenDialog(InformGeneratorApp.primaryStage));

			parser.parse(stream, handler, metadata, parseContext);
			File outFile = new File("output.html");
			// FileOutputStream out = new FileOutputStream(outFile);
			OutputStreamWriter out = new OutputStreamWriter(new FileOutputStream(outFile), StandardCharsets.UTF_8);
			String content = handler.toString();
			// content.replaceFirst("<head a=\"s\">", "<head>");
			//content.replaceFirst("head", "headssss");
			out.write(content);
			out.close();
			
			Document doc = null;
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

			DocumentBuilder builder;
			try {
				builder = factory.newDocumentBuilder();
				doc = builder.parse(new File("output.html"));
				
				NodeList list = doc.getElementsByTagName("head");
				NodeList l = doc.getElementsByTagName("title");
				Element head = (Element) list.item(0);
				head.removeChild(l.item(0)); // Esto es necesario porque por algún motivo (probablemente un bug) DOM se carga la etiqueta de apertura del título
												// (quizás porque está vacío), lo que da a un error en la estructura del documento y a que no se vea en el webview
				Element metaE = doc.createElement("meta");
				metaE.setAttribute("http-equiv", "Content-Type");
				metaE.setAttribute("content", "text/html; charset=utf-8");
				head.appendChild(metaE);
				head.insertBefore(metaE, head.getFirstChild());
				
			} catch (ParserConfigurationException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			//
			DOMSource source = new DOMSource(doc);
		    FileWriter writer = new FileWriter(new File("output.html"));
		    StreamResult result = new StreamResult(writer);

		    TransformerFactory transformerFactory = TransformerFactory.newInstance();
		    Transformer transformer = transformerFactory.newTransformer();
		    transformer.transform(source, result);
			
			
			//System.out.println(content);

			InformGeneratorApp.fileUrl.setValue(outFile.toURI().toURL().toString());

		} catch (

		NullPointerException e) {
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TikaException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TransformerException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@FXML
	void salir(ActionEvent event) {
		ButtonType siButtonType = new ButtonType("Sí", ButtonData.OK_DONE);

		Alert exitAlert = new Alert(AlertType.WARNING, "", siButtonType, ButtonType.CANCEL);
		exitAlert.setTitle("Salir");
		exitAlert.setHeaderText("Está apunto de salir de la aplicación.");
		exitAlert.setContentText("¿Desea salir de la aplicación?");
		exitAlert.initOwner(InformGeneratorApp.primaryStage);

		Optional<ButtonType> result = exitAlert.showAndWait();

		if (result.get() == siButtonType) {
			Stage stage = (Stage) drawerView.getScene().getWindow();
			stage.close();
		}
	}

}
