package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Optional;
import java.util.ResourceBundle;

import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import com.jfoenix.controls.JFXButton;

import integracion.wordseedexporter.InformGeneratorApp;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
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

	private OfficeManager officeManager;
	
	// model
	public ObjectProperty<File> pdfFile = new SimpleObjectProperty<>();

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
		//https://github.com/phip1611/docx4j-search-and-replace-util
		// https://blog.csdn.net/u011781521/article/details/116260048
		// https://jenkov.com/tutorials/javafx/filechooser.html
		
		//Reemplazar texto: https://stackoverflow.com/questions/3391968/text-replacement-in-winword-doc-using-apache-poi
		//https://gist.github.com/aerobium/bf02e443c079c5caec7568e167849dda
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));

		System.out.println("a");
		try {
			File pdfFileOut = new File(Controller.TEMPDOCSFOLDER + File.separator + "output.pdf");
			JodConverter.convert(fileChooser.showOpenDialog(InformGeneratorApp.primaryStage))
					.as(DefaultDocumentFormatRegistry.DOC)
					.to(pdfFileOut)
					.as(DefaultDocumentFormatRegistry.PDF)
					.execute();
			// Esto es para forzar al pdfViewer que cambie de pdf
			pdfFileProperty().set(null);
			pdfFileProperty().set(pdfFileOut);
		} catch (OfficeException e) {
			e.printStackTrace();
		} finally {
			System.out.println("b");
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

	public final ObjectProperty<File> pdfFileProperty() {
		return this.pdfFile;
	}
	

	public final File getPdfFile() {
		return this.pdfFileProperty().get();
	}
	

	public final void setPdfFile(final File pdfFile) {
		this.pdfFileProperty().set(pdfFile);
	}
	

}
