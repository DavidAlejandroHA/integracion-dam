package integracion.informgenerator.controllers;

import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;

import integracion.informgenerator.InformGeneratorApp;
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

	@Override
	public void initialize(URL location, ResourceBundle resources) {

		// load data
		drawerController = new VBoxDrawerController();

		drawerMenu.setSidePane(drawerController.getView());

		// drawerMenu.close();

		// bindings
		
		// TODO: Cambiar InformGeneratorApp.pdfFile por un getter del drawerController
		InformGeneratorApp.pdfFile.addListener((o, ov, nv) -> {
			System.out.println("dsasd\ns");
			pdfViewer.load(nv);
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
