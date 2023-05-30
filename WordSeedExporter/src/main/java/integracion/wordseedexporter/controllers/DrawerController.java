package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;

import org.apache.commons.io.FileUtils;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.office.LocalOfficeManager;

import com.jfoenix.controls.JFXButton;
import com.jfoenix.controls.JFXDrawer;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.model.DocumentManager;
import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ButtonBar.ButtonData;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TextArea;
import javafx.scene.input.MouseEvent;
import javafx.scene.layout.HBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

/**
 * Esta clase representa el controlador de la vista del {@link JFXDrawer} que utiliza la aplicación.
 * @author David Alejandro Hernández Alonso
 *
 */
public class DrawerController implements Initializable {

	// view

	@FXML
	private HBox drawerView;

	@FXML
	private JFXButton importarDocumentoButton;

	@FXML
	private JFXButton importarFuenteButton;

	@FXML
	private JFXButton exportarPdfButton;

	@FXML
	private JFXButton exportarDocumentoButton;

	@FXML
	private JFXButton salirButton;

	private LocalOfficeManager officeManager;

	private JFXDrawer drawerMenu;

	/**
	 * Méodo para inicializar los diferentes elementos necesarios junto al
	 * controlador que los contiene
	 */
	@Override
	public void initialize(URL location, ResourceBundle resources) {
		// bindings
		exportarPdfButton.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			if (Controller.ficheroImportado.get() != null && !Controller.dataSources.get().isEmpty() && Controller.converterReady.get()) {
				return false;
			} else {
				return true;
			}
		}, Controller.ficheroImportado, Controller.dataSources));
		exportarDocumentoButton.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			if (Controller.ficheroImportado.get() != null && !Controller.dataSources.get().isEmpty()) {
				return false;
			} else {
				return true;
			}
		}, Controller.ficheroImportado, Controller.dataSources));
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
		FileChooser fileChooser = new FileChooser();
		fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
		fileChooser.getExtensionFilters().addAll(
				new FileChooser.ExtensionFilter("Microsoft Word Document (2007)", "*.docx"),
				new FileChooser.ExtensionFilter("Microsoft PowerPoint Document (2007)", "*.pptx"),
				new FileChooser.ExtensionFilter("Microsoft Excel Document (2007)", "*.xlsx"),
				new FileChooser.ExtensionFilter("Office Text Document", "*.odt"),
				new FileChooser.ExtensionFilter("Office Presentation Document", "*.odp"),
				new FileChooser.ExtensionFilter("Office SpreadSheet Document", "*.ods"),
				new FileChooser.ExtensionFilter("Office Graphics Document", "*.odg"));
		File f = fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage);
		if (f != null) {
			Controller.ficheroImportado.set(f);
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
		DirectoryChooser chooser = new DirectoryChooser();
		chooser.setInitialDirectory(new File(System.getProperty("user.home")));
		File f = chooser.showDialog(WordSeedExporterApp.primaryStage);
		if (f != null) {
			reemplazarTexto(f, false);
		}
	}

	@FXML
	void onExportarPdf(ActionEvent event) {
		DirectoryChooser chooser = new DirectoryChooser();
		chooser.setInitialDirectory(new File(System.getProperty("user.home")));
		File f = chooser.showDialog(WordSeedExporterApp.primaryStage);
		if (f != null) {
			reemplazarTexto(f, true);
		}
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
	}

	/**
	 * Crea y muestra una alerta advirtiendo al usuario sobre su salida de la
	 * aplicación.<br>
	 * En caso de cancelar su salida, la alerta se cerrará sin cerrar la aplicación
	 */
	public void salirApp() {
		ButtonType siButtonType = new ButtonType("Sí", ButtonData.OK_DONE);
		Alert exitAlert = new Alert(AlertType.WARNING, "", siButtonType, ButtonType.CANCEL);
		Optional<ButtonType> result = Controller.crearAlerta(exitAlert, "Salir",
				"Está apunto de salir de la aplicación.", "¿Desea salir de la aplicación?", true);

		if (result.get() == siButtonType) {
			new Thread(() -> {
				try {
					closeOfficeManager();
				} catch (OfficeException es) {
//					Controller.crearAlerta(AlertType.WARNING, "Advertencia",
//							"No se han podido parar los servicios de LibreOffice/OpenOffice.", null, false);
				}
			}).start();

			deletePdfs();
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
		fileChooser.setInitialDirectory(new File(System.getProperty("user.home")));
		fileChooser.getExtensionFilters().addAll(
				new FileChooser.ExtensionFilter("Microsoft Excel Document (2007)", "*.xlsx"),
				new FileChooser.ExtensionFilter("Office SpreadSheet Document", "*.ods"));
		try {
			docManager.readData(fileChooser.showOpenDialog(WordSeedExporterApp.primaryStage));
		} catch (Exception e) {
			Alert alerta = new Alert(AlertType.ERROR);
			TextArea textoArea = new TextArea();
			textoArea.setText("Error: " + e.getMessage());
			alerta.getDialogPane().setContent(textoArea);
			Controller.crearAlerta(alerta, "Error", "Error al cargar la fuente de datos.",
					// "La fuente de datos contiene un formato incorrecto respecto a la \n" +
					// "gestión de la aplicación."
					null, false);
		}
	}

	/**
	 * Este método se ejecuta al presionar ek botón de "Exportar documento".<br>
	 * Si el archivo a exportar aún no ha sido exportado, no ocurrirá nada.<br>
	 * Si el documento contiene un formato incorrecto o algún tipo de error en la
	 * estructura interna <br>
	 * (p. ej. renombrar el archivo a un formato distinto) se lanzará una alerta
	 * avisando del error ocurrido.
	 * 
	 * @param f La ruta hacia dónde se van a guardar los ficheros generados
	 */
	public void reemplazarTexto(File f, boolean pdf) {
		DocumentManager docManager = new DocumentManager();
		try {
			if (Controller.ficheroImportado.get() != null) { // si no se canceló la importación del documento
				docManager.giveDocument(Controller.ficheroImportado.get(), f, pdf);
			}
		} catch (Exception e) {
			Alert alerta = new Alert(AlertType.ERROR);
			TextArea textoArea = new TextArea();
			textoArea.setText("Error: " + e.getMessage());
			alerta.getDialogPane().setContent(textoArea);
			Controller.crearAlerta(alerta, "Error",
					"Error al procesar el documento. Es posible que el archivo tenga un formato incorrecto\n"
							+ "o se haya eliminado.",
					null, false);
		}
	}

	/**
	 * Cierra los servicios de LibreOffice/OpenOffice con los que la aplicación
	 * trabaja en la creación y conversión<br>
	 * de documentos pdf
	 * 
	 * @throws OfficeException Si ocurre un error al cerrar los servicios de
	 *                         LibreOffice/OpenOffice
	 */
	private void closeOfficeManager() throws OfficeException {
		if (officeManager != null) {
			OfficeUtils.stopQuietly(officeManager);
		}
		// TODO: Hilo de javafx para nueva ventana/alerta indicando que se está cerrando
		// el programa hasta que se cierre
	}

	/**
	 * Elimina todos los archivos pdf de la carpeta de tmpDocs de .WordSeedExporter
	 * (esta carpeta se localiza en la carpeta de usuario)
	 */
	private void deletePdfs() {
		List<File> fileList = new ArrayList<File>(
				FileUtils.listFiles(new File(Controller.TEMPDOCSFOLDER.getPath()), new String[] { "pdf" }, false));
		for (int i = 0; i < fileList.size(); i++) {
			fileList.get(i).delete();
		}
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
	 * Retorna el objeto {@link JFXButton exportarPdfButton} del elemento
	 * DrawerController
	 * 
	 * @return exportarPdfButton
	 */
	public JFXButton getExportarPdfButton() {
		return exportarPdfButton;
	}

}
