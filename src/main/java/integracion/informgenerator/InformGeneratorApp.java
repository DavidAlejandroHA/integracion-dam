package integracion.informgenerator;

import integracion.informgenerator.controllers.Controller;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class InformGeneratorApp extends Application {

	public static Stage primaryStage;
	public Scene escena;

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
