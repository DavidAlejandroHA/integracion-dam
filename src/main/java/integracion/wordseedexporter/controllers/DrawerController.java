package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.Optional;
import java.util.ResourceBundle;

import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.office.LocalOfficeManager;

import com.jfoenix.controls.JFXButton;
import com.jfoenix.controls.JFXDrawer;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.model.DocumentManager;
import javafx.application.Platform;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ButtonType;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class DrawerController implements Initializable {

	// view

	@FXML
	private HBox drawerView;

	@FXML
	private JFXButton importarDocumentoButton;

	@FXML
	private JFXButton importarFuenteButton;

	@FXML
	private JFXButton salirButton;

	private LocalOfficeManager officeManager;

	private JFXDrawer drawerMenu;

	// model
	private ObjectProperty<File> pdfFile = new SimpleObjectProperty<>();

	/**
	 * Méodo para inicializar los diferentes elementos necesarios junto al
	 * controlador que los contiene
	 */
	@Override
	public void initialize(URL location, ResourceBundle resources) {

	}

	/**
	 * Constructor de la clase para cargar la vista del controlador
	 * 
	 * @author David Alejandro Hernández Alonso
	 */
	public DrawerController() {
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/DrawerMenuView.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * Abre un diálogo de importación de ficheros para que el usuario importe
	 * distintos tipos de documentos (.docx, .pptx, .xlsx, .odt, .odp, .ods y .odg)
	 * y pueda gestionar junto a la fuente de datos importada el manejo de la
	 * aplicación.
	 * 
	 * @param event
	 */
	@FXML
	void importarDocumento(ActionEvent event) {
		// https://github.com/phip1611/docx4j-search-and-replace-util
		// https://blog.csdn.net/u011781521/article/details/116260048
		// https://jenkov.com/tutorials/javafx/filechooser.html

		// Reemplazar texto:
		// https://gist.github.com/aerobium/bf02e443c079c5caec7568e167849dda
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));
		fileChooser.getExtensionFilters().addAll(
				new FileChooser.ExtensionFilter("Microsoft Word Document (2007)", "*.docx"),
				new FileChooser.ExtensionFilter("Microsoft PowerPoint Document (2007)", "*.pptx"),
				new FileChooser.ExtensionFilter("Microsoft Excel Document (2007)", "*.xlsx"),
				new FileChooser.ExtensionFilter("Office Text Document", "*.odt"),
				new FileChooser.ExtensionFilter("Office Presentation Document", "*.odp"),
				new FileChooser.ExtensionFilter("Office SpreadSheet Document", "*.ods"),
				new FileChooser.ExtensionFilter("Office Graphics Document", "*.odg"));
		// fileChooser.setInitialDirectory(new File(System.getProperty("user.home") +
		// File.separator + "Desktop"));
		File f = fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage);
		if (f != null) {
			Controller.ficheroImportado.set(f);
		}

	}

	/**
	 * Ejecuta el método {@link #cargarFuente() cargarFuente} para iniciar la
	 * importación de una fuente de datos para la aplicación.
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void importarFuente(ActionEvent event) {
		cargarFuente();
	}

	/**
	 * <p>
	 * Este método se encarga de importar una fuente de datos para la aplicación, de
	 * manera que<br>
	 * solo aceptará archivos xlsx o odt con sus respectivas tablas. Para ello,
	 * ejecuta el método<br>
	 * {@link DocumentManager#readData() readData} junto con el documento entregado
	 * a través de un diálogo de importacion<br>
	 * de documentos.
	 * </p>
	 * <p>
	 * Al invocarlo abrirá una nueva ventana para escoger el fichero de datos a
	 * importar y cargará<br>
	 * sus datos en la aplicación.
	 * </p>
	 * <p>
	 * En caso de fallar a la hora de importar o leer la fuente de datos, sea por un
	 * formato o<br>
	 * renombramiento incorrecto del archivo o por un formato incorrecto en la
	 * lectura de datos<br>
	 * (p. ej. tablas de solo una casilla de altura), se avisará con una alerta del
	 * fallo producido.
	 * </p>
	 */
	public void cargarFuente() {
		DocumentManager docManager = new DocumentManager();
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File("."));
		fileChooser.getExtensionFilters().addAll(
				new FileChooser.ExtensionFilter("Microsoft Excel Document (2007)", "*.xlsx"),
				new FileChooser.ExtensionFilter("Office SpreadSheet Document", "*.ods"));
		try {
			docManager.readData(fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage));
		} catch (Exception e) {
			Alert alerta = new Alert(AlertType.ERROR);
			alerta.setTitle("Error");
			alerta.setHeaderText("Error al cargar la fuente de datos");
			alerta.setContentText(
					"La fuente de datos contiene un formato incorrecto respecto a la \n" + "gestión de la aplicación.");
			e.printStackTrace();
			alerta.initOwner(WordSeedExporterApp.primaryStage);
			alerta.show();
		}
	}

	/**
	 * Este método se ejecuta al presionar ek botón de "Exportar documento".<br>
	 * Si el archivo a exportar aún no ha sido exportado, no ocurrirá nada.<br>
	 * Si el documento contiene un formato incorrecto o algún tipo de error en la
	 * estructura interna <br>
	 * (p. ej. renombrar el archivo a un formato distinto) se lanzará una alerta
	 * avisando del error ocurrido.
	 */
	public void reemplazarTexto() {
		DocumentManager docManager = new DocumentManager();
		try {
			if (Controller.ficheroImportado.get() != null) { // si no se canceló la importación del documento
				docManager.giveDocument(Controller.ficheroImportado.get());
			}
		} catch (Exception e) {
			crearAlerta(AlertType.ERROR, "Error",
					"Error al procesar el documento. Es posible que el archivo tenga un formato incorrecto.",
					"Error: " + e.getMessage(), false);
		}

	}

	/**
	 * Ejecuta el método {@link #reemplazarTexto() reemplazarTexto} para iniciar el
	 * reemplazo de texto en nuevos documentos a través de la fuente de datos y el
	 * archivo previamente importado.
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void onExportarDocumento(ActionEvent event) {
		reemplazarTexto();
	}

	/**
	 * <p>
	 * Invoca una alerta de advertencia de salida de la aplicación en un nuevo hilo
	 * de javafx utilizando el método {@link #salirApp() salirApp}.
	 * </p>
	 * <p>
	 * Se invoca en un nuevo hilo para evitar que los efectos de pulsación de los
	 * botones y otros efectos visuales<br>
	 * se congelen al invocar la alerta
	 * </p>
	 * <p>
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void salir(ActionEvent event) {
		Platform.runLater(() -> {
			salirApp();
		});
		// drawerView.requestFocus();
	}

	/**
	 * Crea y muestra una alerta advirtiendo al usuario sobre su salida de la
	 * aplicación.<br>
	 * En caso de cancelar su salida, la alerta se cerrará sin cerrar la aplicación
	 */
	public void salirApp() {
		ButtonType siButtonType = new ButtonType("Sí", ButtonData.OK_DONE);
		Alert exitAlert = new Alert(AlertType.WARNING, "", siButtonType, ButtonType.CANCEL);
		Optional<ButtonType> result = crearAlerta(exitAlert, "Salir", "Está apunto de salir de la aplicación.",
				"¿Desea salir de la aplicación?", true);

		if (result.get() == siButtonType) {
			Stage stage = (Stage) drawerView.getScene().getWindow();
			stage.close();
		}
	}

	/**
	 * Al presionar el botón de apertura y cerrado del {@link JFXDrawer drawerMenu},
	 * se invoca este método, el cuál gestiona el abrir y cerrar dicho elemento.
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void onDrawerButton(ActionEvent event) {
		if (drawerMenu.isOpening() || drawerMenu.isOpened()) {
			drawerMenu.close();
		} else if (drawerMenu.isClosing() || drawerMenu.isClosed()) {
			drawerMenu.setPrefWidth(600);
			drawerMenu.open();
		}
	}

	/**
	 * Este método se encarga de <b>expandir</b> el tamaño del elemento
	 * {@link JFXDrawer drawerMenu} en caso de que el cursor entre de la zona del
	 * drawerMenu, para poder expandir el menú de opciones.
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void onMouseDrawerEntered(MouseEvent event) {
		if (drawerMenu.isClosed()) { // si está cerrado y no está abriendose
			drawerMenu.setPrefWidth(600);
		}
	}

	/**
	 * Este método se encarga de <b>encoger</b> el tamaño del elemento
	 * {@link JFXDrawer drawerMenu} en caso de que el cursor salga de la zona del
	 * drawerMenu, para poder seleccionar elementos que se sitúan en un orden por
	 * detrás de este (orden de elementos).
	 * 
	 * @param event El ActionEvent al que escucha
	 */
	@FXML
	void onMouseDrawerExited(MouseEvent event) {
		if (!drawerMenu.isPressed() && !drawerMenu.isOpening() && !drawerMenu.isClosing() && !drawerMenu.isOpened()) {
			drawerMenu.setPrefWidth(280);
		}
	}

	/**
	 * Crea una alerta Alert de javafx según el tipo de alerta, el título, la
	 * cabecera y el texto contenido especificado.
	 * 
	 * @param at   El tipo de alerta en cuestión ({@link AlertType AlertType})
	 * @param t    El título de la alerta
	 * @param ht   El texto de cabecera a mostrar
	 * @param ct   El texto de la alerta
	 * @param wait Si los demás procesos de la aplicación se detienen hasta esperar
	 *             una acción en la alerta por el usuario (diálogo bloqueante)
	 */
	private void crearAlerta(AlertType at, String t, String ht, String ct, boolean wait) {
		Alert alerta = new Alert(at);
		alerta.setTitle(t);
		alerta.setHeaderText(ht);
		alerta.setContentText(ct);
		alerta.initOwner(WordSeedExporterApp.primaryStage);
		if (wait) {
			alerta.showAndWait();
		} else {
			alerta.show();
		}
	}

	/**
	 * Modifica el título, el texto de cabecera y de contenido de la alerta ofrecida
	 * y muestra dicha alerta en pantalla.
	 * 
	 * @param al   La alerta a modificar ({@link Alert})
	 * @param t    El título de la alerta
	 * @param ht   El texto de cabecera a mostrar
	 * @param ct   El texto de la alerta
	 * @param wait Si los demás procesos de la aplicación se detienen hasta esperar
	 *             una acción en la alerta por el usuario (diálogo bloqueante)
	 */
	private Optional<ButtonType> crearAlerta(Alert al, String t, String ht, String ct, boolean wait) {
		al.setTitle(t);
		Optional<ButtonType> optional = null;
		al.setHeaderText(ht);
		al.setContentText(ct);
		al.initOwner(WordSeedExporterApp.primaryStage);
		if (wait) {
			optional = al.showAndWait();
		} else {
			al.show();
			optional = Optional.empty();
		}
		return optional;
	}

	/**
	 * Cierra los servicios de LibreOffice/OpenOffice con los que la aplicación
	 * trabaja en la creación y conversión<br>
	 * de documentos pdf
	 * 
	 * @throws OfficeException Si ocurre un error al cerrar los servicios de
	 *                         LibreOffice/OpenOffice
	 */
	public void closeOfficeManager() throws OfficeException {
		if (officeManager != null) {
			OfficeUtils.stopQuietly(officeManager);
		}
		// TODO: Hilo de javafx para nueva ventana/alerta indicando que se está cerrando
		// el programa hasta que se cierre
//		ProgressBar bar = new ProgressBar();
//		bar.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
//		alert.getDialogPane().setContent(bar);
	}

	/**
	 * Retorna la vista del elemento DrawerController
	 * 
	 * @return HBox la vista del Conrolador DrawerController
	 */
	public HBox getView() {
		return drawerView;
	}

	/**
	 * Establece el valor del elemento {@link JFXDrawer drawerMenu}
	 * 
	 * @param drawerMenu El valor ofrecido a establecer
	 */
	public void setDrawerMenu(JFXDrawer drawerMenu) {
		this.drawerMenu = drawerMenu;
	}

	/**
	 * Establece el valor del elemento {@link LocalOfficeManager officeManager}
	 * 
	 * @param officeManager El valor ofrecido a establecer
	 */
	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
	}

	/**
	 * La propiedad {@link ObjectProperty pdfFile} de este elemento
	 */
	public final ObjectProperty<File> pdfFileProperty() {
		return this.pdfFile;
	}

	/**
	 * Retorna el valor de la propiedad {@link ObjectProperty pdfFile} de este
	 * elemento
	 */
	public final File getPdfFile() {
		return this.pdfFileProperty().get();
	}

	/**
	 * Establece un valor a la propiedad {@link ObjectProperty pdfFile}
	 */
	public final void setPdfFile(final File pdfFile) {
		this.pdfFileProperty().set(pdfFile);
	}

}
