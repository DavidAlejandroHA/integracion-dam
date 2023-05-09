package integracion.informgenerator.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;

import integracion.wordseedexporter.InformGeneratorApp;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.layout.AnchorPane;

public class Controller implements Initializable {

	@FXML
	private AnchorPane view;

	@FXML
	private JFXDrawer drawerMenu;

	@FXML
	private PDFView pdfViewer;

	private VBoxDrawerController drawerController;

	// paths y nombres de rutas
	
	public static final String APPFOLDERNAME = "WordSeedExporter";

	public static final File APPFOLDER = new File(
			System.getProperty("user.home") + File.separator + "." + APPFOLDERNAME);
	
	public static final File TEMPDOCSFOLDER = new File(APPFOLDER.getPath() + File.separator + "tmpDocs");

	@Override
	public void initialize(URL location, ResourceBundle resources) {

		// load data
		drawerController = new VBoxDrawerController();

		drawerMenu.setSidePane(drawerController.getView());

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

		InformGeneratorApp.primaryStage.setOnCloseRequest(e -> {
			drawerController.closeOfficeManager();
			System.out.println("assdds2");
			// TODO: Iniciar nuevo hilo javafx y ejecutar ahí el closeOfficeManager()
		});
	}

	// Hacer otro controlador con la vista del drawerView y ponerle el getView()

	public Controller() {
		// https://stackoverflow.com/questions/13815119/apache-poi-converting-doc-to-html-with-images
		// https://stackoverflow.com/questions/7868713/convert-word-to-html-with-apache-poi
		// https://poi.apache.org/components/
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/View.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public AnchorPane getView() {
		return view;
	}

}
