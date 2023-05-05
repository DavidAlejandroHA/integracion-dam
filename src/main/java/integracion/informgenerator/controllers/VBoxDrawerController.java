package integracion.informgenerator.controllers;

import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import com.jfoenix.controls.JFXButton;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;

public class VBoxDrawerController implements Initializable {
	
	@FXML
    private HBox drawerView;

    @FXML
    private JFXButton importarDocumentoButton;
    
    @FXML
    private JFXButton am;

	@Override
	public void initialize(URL location, ResourceBundle resources) {
		// TODO Auto-generated method stub

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

}
