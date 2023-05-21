package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.office.LocalOfficeManager;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;
import com.jfoenix.controls.events.JFXDrawerEvent;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.components.PDFViewSkinES;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.ListProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleListProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.layout.AnchorPane;

public class Controller implements Initializable {

	@FXML
	private AnchorPane view;

	@FXML
	private JFXDrawer drawerMenu;

	@FXML
	private PDFView pdfViewer;

	private VBoxDrawerController drawerController;

	private LocalOfficeManager officeManager;

	// paths y nombres de rutas

	public static final String APPFOLDERNAME = "WordSeedExporter";

	public static final File APPFOLDER = new File(
			System.getProperty("user.home") + File.separator + "." + APPFOLDERNAME);

	public static final File TEMPDOCSFOLDER = new File(APPFOLDER.getPath() + File.separator + "tmpDocs");

	// model
	public static BooleanProperty replaceExactWord = new SimpleBooleanProperty(true); // valor por defecto a true

	// private ObservableList<String> replaceList =
	// FXCollections.observableArrayList();

	public static ListProperty<ObservableList<String>> columnList = new SimpleListProperty<>(
			FXCollections.<ObservableList<String>>observableArrayList());
	// https://stackoverflow.com/questions/16317949/javafx-two-dimensional-observablelist

	// La listproperty dentro de columnList deberá iterar por el mismo índice que
	// la listproperty keyList
	public static ListProperty<String> keyList = new SimpleListProperty<>(FXCollections.observableArrayList());

	@Override
	public void initialize(URL location, ResourceBundle resources) {

		// load data
		ObservableList<String> replaceList = FXCollections.observableArrayList("soso", "toro");
		Controller.columnList.add(replaceList);
		Controller.keyList.add("pepe");
//		columnList TODO Continuar Todo
		// replaceList.setAll(FXCollections.observableArrayList());
		// columnList.add(replaceList);
		drawerController = new VBoxDrawerController();
		drawerMenu.setSidePane(drawerController.getView());
		drawerMenu.setPrefWidth(280);
		drawerController.setDrawerMenu(drawerMenu);

		// drawerMenu.close();
		// create app folder

		if (!Controller.APPFOLDER.exists()) {
			Controller.APPFOLDER.mkdirs();
			Controller.TEMPDOCSFOLDER.mkdirs();
		}

		// listeners

		drawerController.pdfFileProperty().addListener((o, ov, nv) -> {
			// El if y el else es para forzar al pdfViewer que cambie de pdf
			if (nv != null) {
				pdfViewer.load(nv);
			} else {
				pdfViewer.unload();
			}
		});

		// Añadir la interfaz personalizada en Español al pdfViewer
		pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));

		// Añadir hilo para cerrar
		WordSeedExporterApp.primaryStage.setOnCloseRequest(e -> {
			try {
				drawerController.closeOfficeManager();
			} catch (OfficeException es) {
				es.printStackTrace();
			}
		});
	}

	public Controller() {
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/View.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@FXML
	void onDrawerClosed(JFXDrawerEvent event) {
		drawerMenu.setPrefWidth(300);
	}

	@FXML
	void onDrawerOpened(JFXDrawerEvent event) {
		// drawerMenu.setPrefWidth(300);
	}

	public AnchorPane getView() {
		return view;
	}

	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
		try {
			this.officeManager.start();
			drawerController.setOfficeManager(this.officeManager);
		} catch (OfficeException e) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error");
			alert.setHeaderText("Se ha producido un error al iniciar LibeOffice/OpenOffice");
			alert.initOwner(WordSeedExporterApp.primaryStage);
			alert.show();
		}
	}
}
