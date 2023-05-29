package integracion.wordseedexporter;

import org.jodconverter.local.office.LocalOfficeManager;

import integracion.wordseedexporter.controllers.Controller;
import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

public class WordSeedExporterApp extends Application {

	public static Stage primaryStage;
	public static Scene escena;
	private Controller controller;
	private LocalOfficeManager officeManager;

	/**
	 * El LocalOfficeManager se define e inicializa en esta clase por los siguientes
	 * motivos:
	 * 
	 * 1: Permite lanzar una alerta que se lanza en primer plano (que aparezca
	 * delante de la escena principal) nada más iniciar la aplicación y que use el
	 * icono de la aplicación (ambas partes gracias al método .initOwner)
	 * 
	 * 2: Permite contener un enlace funcional que al hacer click te lleva a la
	 * página de libreoffice.
	 * 
	 * El método getHostServices() sólo está disponible para clases que extienden de
	 * Application, por lo que no es posible utilizarlo en los controladores a no
	 * ser que reciban una instancia desde la aplicación principal.
	 */

	@Override
	public void start(Stage primaryStage) throws Exception {
		WordSeedExporterApp.primaryStage = primaryStage;

		// para gestionar que primaryStage no sea nulo
		controller = new Controller();
		escena = new Scene(controller.getView());

		primaryStage.setScene(escena);
		primaryStage.setTitle("WordSeedExporter");
		// primaryStage.getIcons().add(new
		// Image(WordSeedExporterApp.class.getResourceAsStream("/images/calendar-16x16.png")));
		primaryStage.show();

		/**
		 * Este método se ejecuta después de todos los demás para tener referencias al
		 * primaryStage e inicializar otros elementos
		 */
		checkOffice();
	}

	/**
	 * Método que establece un valor al LocalOfficeManager del controlador principal
	 * para la visualización de los documentos importados.
	 * 
	 * En caso de no estar instalado el LibreOffice o el OpenOffice, saldrá una
	 * alerta avisando de la situación
	 */
	public void checkOffice() {
		try {
			officeManager = LocalOfficeManager.install();
			controller.setOfficeManager(officeManager);
		} catch (NullPointerException e) {
			Controller.officeInstalled = false;
			Alert nullPointExAlert = new Alert(AlertType.WARNING);
			nullPointExAlert.setTitle("Office no está instalado");
			nullPointExAlert.setHeaderText("LibreOffice/Openoffice no está instalado en este equipo.");
			VBox alertContent = new VBox();

			Hyperlink link = new Hyperlink("aquí");
			link.setOnAction(t -> {
				this.getHostServices().showDocument("https://www.libreoffice.org/download/download-libreoffice/");
			});
			Label l1 = new Label("Instale LibreOffice u OpenOffice en su equipo para poder efectuar");
			Label l2 = new Label("la previsualización y exportación de documentos a pdf.");
			Label l3 = new Label("Haga click");
			link.setText("aquí");
			Label l4 = new Label(" si desea instalar LibreOffice en su equipo.");
			HBox l3Content = new HBox(l3, link, l4);
			link.setAlignment(Pos.TOP_CENTER);
			alertContent.getChildren().addAll(l1, l2, l3Content);
			
			nullPointExAlert.getDialogPane().setContent(alertContent);
			nullPointExAlert.initOwner(WordSeedExporterApp.primaryStage);
			nullPointExAlert.show();
		}

	}
}
