package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.office.LocalOfficeManager;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.components.PDFViewSkinES;
import javafx.concurrent.Task;
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

	private LocalOfficeManager officeManager;

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
		
		// Añadir la interfaz personalizada en Español al pdfViewer
		pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));

		WordSeedExporterApp.primaryStage.setOnCloseRequest(e -> {
			// Iniciando nuevo hilo javafx y ejecutar ahí el closeOfficeManager()
			Task<Void> task = new Task<>() {
				protected Void call() throws Exception {
					drawerController.closeOfficeManager();
					System.out.println("closed");
					return null;
				}
			};
			new Thread(task).start();
		});
	}

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
	
	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
		try {
			this.officeManager.start();
		} catch (OfficeException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
