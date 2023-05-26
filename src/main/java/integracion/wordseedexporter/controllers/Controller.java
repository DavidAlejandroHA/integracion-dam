package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Optional;
import java.util.ResourceBundle;

import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;
import com.jfoenix.controls.events.JFXDrawerEvent;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.components.PDFViewSkinES;
import integracion.wordseedexporter.model.DataSource;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.ListProperty;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleListProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.layout.AnchorPane;

public class Controller implements Initializable {

	@FXML
	private AnchorPane view;

	@FXML
	private JFXDrawer drawerMenu;

	@FXML
	private PDFView pdfViewer;

	private DrawerController drawerController;

	private LocalOfficeManager officeManager;

	// paths y nombres de rutas

	public static final String APPFOLDERNAME = "WordSeedExporter";

	public static final File APPFOLDER = new File(
			System.getProperty("user.home") + File.separator + "." + APPFOLDERNAME);

	public static final File TEMPDOCSFOLDER = new File(APPFOLDER.getPath() + File.separator + "tmpDocs");

	// model
	public static BooleanProperty replaceExactWord = new SimpleBooleanProperty(true); // valor por defecto a true

	//public static ListProperty<ObservableList<String>> rowList = new SimpleListProperty<>(
			//FXCollections.<ObservableList<String>>observableArrayList());
	// https://stackoverflow.com/questions/16317949/javafx-two-dimensional-observablelist

	//public static ListProperty<String> keyList = new SimpleListProperty<>(FXCollections.observableArrayList());
	
	public static ListProperty<DataSource> dataSources = new SimpleListProperty<>(FXCollections.observableArrayList());

	public static ObjectProperty<File> ficheroImportado = new SimpleObjectProperty<>();

	@Override
	public void initialize(URL location, ResourceBundle resources) {

		// load data
		drawerController = new DrawerController();
		drawerMenu.setSidePane(drawerController.getView());
		drawerMenu.setPrefWidth(280);
		drawerController.setDrawerMenu(drawerMenu);
		pdfViewer.setDisable(true);

		// crear carpeta de la app si no existe

		if (!Controller.APPFOLDER.exists()) {
			Controller.APPFOLDER.mkdirs();
			Controller.TEMPDOCSFOLDER.mkdirs();
		}

		// listeners

		drawerController.pdfFileProperty().addListener((o, ov, nv) -> {
			// El if y el else es para forzar al pdfViewer que cambie de pdf
			if (nv != null) {
				pdfViewer.setDisable(false);
				pdfViewer.load(nv);
			} else {
				pdfViewer.setDisable(true);
				pdfViewer.unload();
			}
		});

		// listener para el fichero importado
		// en este caso se intenta hacer una previsualización del documento en el visor
		// de pdfs
		ficheroImportado.addListener((o, ov, nv) -> {
			if (nv != null) {
				try {
					File pdfFileOut = new File(Controller.TEMPDOCSFOLDER + File.separator + "preview.pdf");
					JodConverter.convert(nv)// .as(DefaultDocumentFormatRegistry.DOC)
							.to(pdfFileOut)//
							.as(DefaultDocumentFormatRegistry.PDF)//
							.execute();
					// El null es para forzar al pdfViewer que cambie de pdf
					
					drawerController.pdfFileProperty().set(null);
					drawerController.pdfFileProperty().set(pdfFileOut);

				} catch (OfficeException | IllegalStateException e) { // informar de que
					crearAlerta(AlertType.WARNING, "Advertencia",
							"No se ha podido crear la previsualización \n" + "del documento importado.", null, false);
				}
				System.out.println(nv);
			}
		});

		// Añadir la interfaz personalizada en Español al pdfViewer
		pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));

		// Crear setOnCloseRequest para cerrar los servicios de office y advertir de
		// salida
		WordSeedExporterApp.primaryStage.setOnCloseRequest(e -> {
			try {
				drawerController.closeOfficeManager();
			} catch (OfficeException es) {
				crearAlerta(AlertType.WARNING, "Advertencia",
						"No se han podido parar los servicios de LibreOffice/OpenOffice.", null, false);
			}
			drawerController.salirApp();
			e.consume(); // Si llega hasta aquí es porque el usuario ha decidido cancelar la salida de la
							// aplicación
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
		drawerMenu.setPrefWidth(280);
	}

	@FXML
	void onDrawerOpened(JFXDrawerEvent event) {
		// drawerMenu.setPrefWidth(600);
	}

	/**
	 * Retorna la vista del elemento DrawerController
	 * 
	 * @return AnchorPane la vista del Conrolador Principal
	 */
	public AnchorPane getView() {
		return view;
	}

	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
		try {
			this.officeManager.start();
			drawerController.setOfficeManager(this.officeManager);
		} catch (OfficeException e) {
			crearAlerta(AlertType.ERROR, "Error", "Se ha producido un error al iniciar LibeOffice/OpenOffice", null,
					false);
		}
	}

	/**
	 * Crea una alerta Alert de javafx según el tipo de alerta, el título, la
	 * cabecera y el texto contenido especificado.
	 * 
	 * @param at   El tipo de alerta en cuestión ({@link AlertType AlertType})
	 * @param t    El título de la alerta
	 * @param ht   El texto de cabecera a mostrar
	 * @param ct   El texto de la alerta
	 * @param wait Si los demás procesos de la aplicación se detienen hasta esperar
	 *             una acción en la alerta por el usuario (diálogo bloqueante)
	 */
	public static void crearAlerta(AlertType at, String t, String ht, String ct, boolean wait) {
		Alert alerta = new Alert(at);
		alerta.setTitle(t);
		alerta.setHeaderText(ht);
		alerta.setContentText(ct);
		alerta.initOwner(WordSeedExporterApp.primaryStage);
		if (wait) {
			alerta.showAndWait();
		} else {
			alerta.show();
		}
	}

	/**
	 * Modifica el título, el texto de cabecera y de contenido de la alerta ofrecida
	 * y muestra dicha alerta en pantalla.
	 * 
	 * @param al   La alerta a modificar ({@link Alert})
	 * @param t    El título de la alerta
	 * @param ht   El texto de cabecera a mostrar
	 * @param ct   El texto de la alerta
	 * @param wait Si los demás procesos de la aplicación se detienen hasta esperar
	 *             una acción en la alerta por el usuario (diálogo bloqueante)
	 */
	public static Optional<ButtonType> crearAlerta(Alert al, String t, String ht, String ct, boolean wait) {
		al.setTitle(t);
		Optional<ButtonType> optional = null;
		al.setHeaderText(ht);
		al.setContentText(ct);
		al.initOwner(WordSeedExporterApp.primaryStage);
		if (wait) {
			optional = al.showAndWait();
		} else {
			al.show();
			optional = Optional.empty();
		}
		return optional;
	}
}
