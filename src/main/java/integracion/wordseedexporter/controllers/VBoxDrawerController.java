package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Optional;
import java.util.ResourceBundle;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import com.jfoenix.controls.JFXButton;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.model.DocumentManager;
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
	private JFXButton importarFuenteButton;

	private LocalOfficeManager officeManager;

	// model
	public ObjectProperty<File> pdfFile = new SimpleObjectProperty<>();

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

	public void closeOfficeManager() {
		if (officeManager != null) {
			OfficeUtils.stopQuietly(officeManager);
		}
		// TODO: Hilo de javafx para nueva ventana/alerta indicando que se está cerrando
		// el programa hasta que se cierre
	}

	@FXML
	void importarDocumento(ActionEvent event) {
		// https://github.com/phip1611/docx4j-search-and-replace-util
		// https://blog.csdn.net/u011781521/article/details/116260048
		// https://jenkov.com/tutorials/javafx/filechooser.html

		// Reemplazar texto:
		// https://gist.github.com/aerobium/bf02e443c079c5caec7568e167849dda
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));

		try {
			File pdfFileOut = new File(Controller.TEMPDOCSFOLDER + File.separator + "output.pdf");
			File outDir = fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage);
			if (outDir != null) {
				JodConverter.convert(outDir).as(DefaultDocumentFormatRegistry.DOC).to(pdfFileOut)
						.as(DefaultDocumentFormatRegistry.PDF).execute();
			}

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
	void importarFuente(ActionEvent event) {
		reemplazarTexto();
	}

	public void reemplazarTexto() {
		DocumentManager docManager = new DocumentManager();
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));
		fileChooser.getExtensionFilters().addAll(//
				new FileChooser.ExtensionFilter("Microsoft Word Document (2007)", "*.docx"),
				new FileChooser.ExtensionFilter("Microsoft PowerPoint Document (2007)", "*.pptx"),
				new FileChooser.ExtensionFilter("Microsoft Excel Document (2007)", "*.xlsx"),
				new FileChooser.ExtensionFilter("Office Text Document", "*.odt"),
				new FileChooser.ExtensionFilter("Office Presentation Document", "*.odp"));
		try {
			docManager.giveDocument(fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage));
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// TODO: Terminar la importación de la fuente de datos
	}

	@FXML
	void salir(ActionEvent event) {
		ButtonType siButtonType = new ButtonType("Sí", ButtonData.OK_DONE);

		Alert exitAlert = new Alert(AlertType.WARNING, "", siButtonType, ButtonType.CANCEL);
		exitAlert.setTitle("Salir");
		exitAlert.setHeaderText("Está apunto de salir de la aplicación.");
		exitAlert.setContentText("¿Desea salir de la aplicación?");
		exitAlert.initOwner(WordSeedExporterApp.primaryStage);

		Optional<ButtonType> result = exitAlert.showAndWait();

		if (result.get() == siButtonType) {
			Stage stage = (Stage) drawerView.getScene().getWindow();
			stage.close();
		}
	}

	// TODO
	@FXML
	void enterMouseOnDrawer(MouseEvent event) {
		System.out.println("a");
	}

	@FXML
	void exitMouseOfDrawer(MouseEvent event) {
		System.out.println("b");
	}

	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
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
