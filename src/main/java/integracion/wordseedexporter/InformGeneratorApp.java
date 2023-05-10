package integracion.wordseedexporter;

import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.office.LocalOfficeManager;

import integracion.wordseedexporter.controllers.Controller;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

public class InformGeneratorApp extends Application {

	public static Stage primaryStage;
	public Scene escena;
	private Controller controller;
	private LocalOfficeManager officeManager;

	@Override
	public void start(Stage primaryStage) throws Exception {
		InformGeneratorApp.primaryStage = primaryStage;

		// para gestionar que primaryStage no sea nulo
		controller = new Controller();
		escena = new Scene(controller.getView());

		primaryStage.setScene(escena);
		primaryStage.setTitle("InformGenerator");
		// primaryStage.getIcons().add(new
		// Image(CalendarApp.class.getResourceAsStream("/images/calendar-16x16.png")));
		primaryStage.show();

		checkOffice();
	}

	public void checkOffice() {
		try {
			officeManager = LocalOfficeManager.install();
			controller.setOfficeManager(officeManager);
		} catch (NullPointerException e) {
			Alert nullPointExAlert = new Alert(AlertType.WARNING);
			nullPointExAlert.setTitle("LibreOffice no está instalado");
			nullPointExAlert.setHeaderText("LibreOffice no está instalado en este equipo.");
			VBox alertContent = new VBox();

			Hyperlink link = new Hyperlink("aquí");
			link.lineSpacingProperty().set(10.0);
			link.setPrefHeight(10.0);
			link.setMaxHeight(10.0);
			link.setOnAction(t -> {
				this.getHostServices().showDocument("https://www.libreoffice.org/download/download-libreoffice/");
			});
			Label l1 = new Label("Instale LibreOffice u OpenOffice en su equipo para poder ver");
			Label l2 = new Label("los documentos importados en la aplicación. Haga click ");
			link.setText("aquí");
			Label l3 = new Label("si desea instalar LibreOffice en su equipo.");
			HBox l2Content = new HBox(l2, link);
			
			alertContent.getChildren().addAll(l1, l2Content, l3);
			nullPointExAlert.getDialogPane().contentProperty().set(alertContent);
			nullPointExAlert.initOwner(InformGeneratorApp.primaryStage);
			nullPointExAlert.show();
		}

	}

	public static void main(String[] args) {
		launch(args);
	}

}
