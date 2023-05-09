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
import java.nio.file.Files;
import java.nio.file.Path;
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
import org.apache.tika.metadata.TikaCoreProperties;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.ocr.TesseractOCRConfig;
import org.apache.tika.parser.pdf.PDFParserConfig;
import org.apache.tika.sax.ToXMLContentHandler;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.apache.tika.extractor.EmbeddedDocumentExtractor;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.ContentHandler;
import org.xml.sax.SAXException;

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

	private LocalOfficeManager officeManager;
	// model

	WordToHtmlConverter wordToHtmlConverter;
	HWPFDocumentCore wordDocument;

	@Override
	public void initialize(URL location, ResourceBundle resources) {
		try {
			officeManager = LocalOfficeManager.install();
			officeManager.start();
			//NullPointerException: officeHome must not be null
		} catch (NullPointerException e) {
			// TODO Hacer alerta
			e.printStackTrace();
			Alert nullPointExAlert = new Alert(AlertType.WARNING);
			nullPointExAlert.setTitle(null);
		} catch (OfficeException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
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

	public void closeOfficeManager() {
		OfficeUtils.stopQuietly(officeManager);
	}

	@FXML
	void importarDocumento(ActionEvent event) {
		//Docx4JSRUtil.searchAndReplace
		//https://github.com/phip1611/docx4j-search-and-replace-util
		// https://blog.csdn.net/u011781521/article/details/116260048
		// https://jenkov.com/tutorials/javafx/filechooser.html
		
		//Reemplazar texto: https://stackoverflow.com/questions/3391968/text-replacement-in-winword-doc-using-apache-poi
		//https://gist.github.com/aerobium/bf02e443c079c5caec7568e167849dda
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));

		System.out.println("a");
		try {
			File pdfFile = new File("output.pdf");
			JodConverter.convert(fileChooser.showOpenDialog(InformGeneratorApp.primaryStage))
					.as(DefaultDocumentFormatRegistry.DOC)
					.to(pdfFile)
					.as(DefaultDocumentFormatRegistry.PDF)
					.execute();
			InformGeneratorApp.pdfFile.set(pdfFile);
		} catch (OfficeException e) {
			e.printStackTrace();
		} finally {
			System.out.println("b");
		}
	}

	@FXML
	void importarDocumento_(ActionEvent event) {
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));
		
		
		// https://itecnote.com/tecnote/extract-images-from-pdf-with-apache-tika/
		/*
		 * try { System.out.println(FileMagic.valueOf(fileChooser.showOpenDialog(
		 * InformGeneratorApp.primaryStage))); } catch (IOException e1) { // TODO
		 * Auto-generated catch block e1.printStackTrace(); }
		 */
		try {

			AutoDetectParser parser = new AutoDetectParser();
			Metadata metadata = new Metadata();
			ContentHandler handler = new ToXMLContentHandler();

			TesseractOCRConfig config = new TesseractOCRConfig();
			PDFParserConfig pdfConfig = new PDFParserConfig();
			pdfConfig.setExtractInlineImages(true);

			// To parse images in files those lines are needed
			ParseContext parseContext = new ParseContext();
			parseContext.set(TesseractOCRConfig.class, config);
			parseContext.set(PDFParserConfig.class, pdfConfig);
			parseContext.set(Parser.class, parser); // need to add this to make sure
													// recursive parsing happens!
			// parseContext.set

			EmbeddedDocumentExtractor embeddedDocumentExtractor = new EmbeddedDocumentExtractor() {
				@Override
				public boolean shouldParseEmbedded(Metadata metadata) {
					return false;
				}

				@Override
				public void parseEmbedded(InputStream stream, ContentHandler handler, Metadata metadata,
						boolean outputHtml) throws SAXException, IOException {
					Path outputFile = new File(metadata.get(TikaCoreProperties.RESOURCE_NAME_KEY)).toPath();
					Files.copy(stream, outputFile);
				}
			};

			parseContext.set(EmbeddedDocumentExtractor.class, embeddedDocumentExtractor);

			InputStream stream = new FileInputStream(fileChooser.showOpenDialog(InformGeneratorApp.primaryStage));

			parser.parse(stream, handler, metadata, parseContext);
			File outFile = new File("output.html");
			OutputStreamWriter out = new OutputStreamWriter(new FileOutputStream(outFile), StandardCharsets.UTF_8);
			String content = handler.toString();
			out.write(content);
			out.close();

			Document doc = null;
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();

			DocumentBuilder builder ;
			try {
				builder = factory.newDocumentBuilder();
				doc = builder.parse(new File("output.html"));

				NodeList list = doc.getElementsByTagName("head");
				NodeList listTitle = doc.getElementsByTagName("title");
				NodeList listImages = doc.getElementsByTagName("img");
				NodeList listBody = doc.getElementsByTagName("body");
				// <style>p:input:not([value=""]) { margin: 0; }</style>

				Element head = (Element) list.item(0);
				head.removeChild(listTitle.item(0)); // Esto es necesario porque por algún motivo (probablemente un bug)
														// DOM se carga la etiqueta de apertura del título
				// (quizás porque está vacío), lo que da a un error en la estructura del
				// documento y a que no se vea en el webview
				Element metaE = doc.createElement("meta");
				metaE.setAttribute("http-equiv", "Content-Type");
				metaE.setAttribute("content", "text/html; charset=utf-8");
				head.appendChild(metaE);
				head.insertBefore(metaE, head.getFirstChild());

				// Remover el "embedded:" del atributo de src
				for (int i = 0; i < listImages.getLength(); i++) {
					String newAttr = ((Element) listImages.item(i)).getAttribute("src");
					newAttr = newAttr.replaceFirst("embedded:", "");
					((Element) listImages.item(i)).setAttribute("src", newAttr);
				}

				Element style = doc.createElement("style");
				style.setTextContent("p:input:not([value=\"\"]) { margin: 0; }");

				Element body = ((Element) listBody.item(0));
				body.appendChild(style);
				body.insertBefore(style, body.getFirstChild());

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

			// System.out.println(content);

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
