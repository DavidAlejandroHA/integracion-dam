package integracion.informgenerator.controllers;

import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import com.jfoenix.controls.JFXDrawer;

import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.layout.AnchorPane;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;

public class Controller implements Initializable{

	@FXML
	private AnchorPane view;
	
	@FXML
	private WebView webView;
	
	@FXML
    private JFXDrawer drawerMenu;
	
	private WebEngine webEngine;
	
	private VBoxDrawerController drawerController;
	
	
	
	@Override
	public void initialize(URL location, ResourceBundle resources) {
		drawerController = new VBoxDrawerController();
		webEngine = webView.getEngine();
		webEngine.load("https://www.google.es");
		drawerMenu.setSidePane(drawerController.getView());
		
		
		drawerMenu.close();
		//drawerMenu.setVisible(true);
		
	}
	
	// Hacer otro controlador con la vista del drawerView y ponerle el getView()
	
	public Controller() {
		//https://stackoverflow.com/questions/13815119/apache-poi-converting-doc-to-html-with-images
		//https://stackoverflow.com/questions/7868713/convert-word-to-html-with-apache-poi
		//https://poi.apache.org/components/
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
