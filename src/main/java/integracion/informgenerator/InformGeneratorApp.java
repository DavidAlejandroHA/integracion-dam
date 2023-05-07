package integracion.informgenerator;

import java.io.File;

import integracion.informgenerator.controllers.Controller;
import javafx.application.Application;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class InformGeneratorApp extends Application {

	public static Stage primaryStage;
	public Scene escena;
	public static StringProperty webViewContent = new SimpleStringProperty();
	public static StringProperty fileUrl =  new SimpleStringProperty();

	private Controller controller = new Controller();

	@Override
	public void start(Stage primaryStage) throws Exception {
		primaryStage.setTitle("InformGenerator");
		escena = new Scene(controller.getView());
		primaryStage.setScene(escena);
		// primaryStage.getIcons().add(new
		// Image(CalendarApp.class.getResourceAsStream("/images/calendar-16x16.png")));
		InformGeneratorApp.primaryStage = primaryStage;
		primaryStage.show();
	}

	public static void main(String[] args) {
		launch(args);
	}

}
