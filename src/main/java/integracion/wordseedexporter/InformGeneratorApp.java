package integracion.wordseedexporter;

import java.io.File;

import integracion.wordseedexporter.controllers.Controller;
import javafx.application.Application;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class InformGeneratorApp extends Application {

	public static Stage primaryStage;
	public Scene escena;
	public static ObjectProperty<File> pdfFile = new SimpleObjectProperty<>();
	public static StringProperty fileUrl =  new SimpleStringProperty();

	private Controller controller;

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
	}

	public static void main(String[] args) {
		launch(args);
	}

}
