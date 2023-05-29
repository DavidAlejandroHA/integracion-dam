package integracion.wordseedexporter.controllers;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Optional;
import java.util.ResourceBundle;
import java.util.function.UnaryOperator;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.jodconverter.core.document.DefaultDocumentFormatRegistry;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.local.JodConverter;
import org.jodconverter.local.office.LocalOfficeManager;

import com.dlsc.pdfviewfx.PDFView;
import com.jfoenix.controls.JFXDrawer;
import com.jfoenix.controls.events.JFXDrawerEvent;

import integracion.wordseedexporter.WordSeedExporterApp;
import integracion.wordseedexporter.components.PDFViewSkinES;
import integracion.wordseedexporter.model.DataSource;
import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.beans.binding.BooleanBinding;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.ListProperty;
import javafx.beans.property.ObjectProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleListProperty;
import javafx.beans.property.SimpleObjectProperty;
import javafx.collections.FXCollections;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.CheckBox;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.TextFormatter;
import javafx.scene.control.TextFormatter.Change;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;

public class Controller implements Initializable {

	@FXML
	private AnchorPane view;

	@FXML
	private JFXDrawer drawerMenu;

	@FXML
	private HBox loadingBackground;

	@FXML
	private PDFView pdfViewer;

	@FXML
	private Button leftButton;

	@FXML
	private Button rightButton;

	@FXML
	private Button topLeftButton;

	@FXML
	private Button topRightButton;

	@FXML
	private Button unloadButton;

	@FXML
	private Button reloadButton;

	@FXML
	private CheckBox distinctMayusCheckBox;

	@FXML
	private CheckBox exactWordsCheckBox;

	@FXML
	private TextField pageNumTextField;

	private DrawerController drawerController;

	private LocalOfficeManager officeManager;

	// paths y nombres de rutas

	public static final String APPFOLDERNAME = "WordSeedExporter";

	public static final File APPFOLDER = new File(
			System.getProperty("user.home") + File.separator + "." + APPFOLDERNAME);

	public static final File TEMPDOCSFOLDER = new File(APPFOLDER.getPath() + File.separator + "tmpDocs");

	// model
	public static BooleanProperty replaceExactWord = new SimpleBooleanProperty(true); // valor por defecto a true
																						// (wordBoundaries)

	public static BooleanProperty caseSensitive = new SimpleBooleanProperty(true); // valor por defecto a true

	public static ListProperty<DataSource> dataSources = new SimpleListProperty<>(FXCollections.observableArrayList());

	public static ObjectProperty<File> ficheroImportado = new SimpleObjectProperty<>();

	public static ListProperty<File> previsualizaciones = new SimpleListProperty<>(FXCollections.observableArrayList());

	private static ObjectProperty<File> pdfCreado = new SimpleObjectProperty<>();

	private static BooleanProperty previewReady = new SimpleBooleanProperty(false);

	private static IntegerProperty documentIndex = new SimpleIntegerProperty(1);

	public static boolean officeInstalled = true; // valor por defecto

	/**
	 * Inicia los objetos y métodos necesarios para cargar todos los elementos e
	 * interfaces necesarios de la aplicación
	 */
	@Override
	public void initialize(URL location, ResourceBundle resources) {

		// Quitar advertencias del logger
		org.apache.log4j.Logger.getRootLogger().setLevel(org.apache.log4j.Level.INFO);
		// load data
		drawerController = new DrawerController();
		drawerMenu.setSidePane(drawerController.getView());
		drawerMenu.setPrefWidth(280);
		drawerController.setDrawerMenu(drawerMenu);
		pdfViewer.setDisable(true);

		// crear carpeta de la app si no existe
		if (!Controller.APPFOLDER.exists()) {
			Controller.APPFOLDER.mkdirs();
			Controller.TEMPDOCSFOLDER.mkdirs();
		}

		// listeners
		pdfCreado.addListener((o, ov, nv) -> {
			// En ciertas ocasiones, al volver a cargar un nuevo documento el visor de pdfs
			// lanza una exepción desde uno de sus hilos (al cargar documentos con páginas
			// en horizontal) que no afecta a la aplicación en sí, pero para evitarlo se
			// inicia un nuedo visor de pdfs
			pdfViewer = new PDFView();
			pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));
			pdfViewer.setMinWidth(Region.USE_COMPUTED_SIZE);
			pdfViewer.setPrefWidth(Region.USE_COMPUTED_SIZE);
			pdfViewer.setMaxWidth(Region.USE_COMPUTED_SIZE);
			pdfViewer.setMinHeight(Region.USE_COMPUTED_SIZE);
			pdfViewer.setMaxHeight(Region.USE_COMPUTED_SIZE);
			pdfViewer.setPrefHeight(Region.USE_COMPUTED_SIZE);
			HBox.setHgrow(pdfViewer, Priority.ALWAYS);
			pdfViewer.load(nv);
			HBox hb = (HBox) view.lookup("#pdfHBox");
			hb.getChildren().set(1, pdfViewer);
		});

		// listener para el fichero importado
		// en este caso se intenta hacer una previsualización del documento en el visor
		// de pdfs
		ficheroImportado.addListener((o, ov, nv) -> {
			if (nv != null && previewReady.get()) {
				File pdfFileOut = new File(Controller.TEMPDOCSFOLDER + File.separator + "preview.pdf");
				try {
					JodConverter.convert(nv)// .as(DefaultDocumentFormatRegistry.DOC)
							.to(pdfFileOut)//
							.as(DefaultDocumentFormatRegistry.PDF)//
							.execute();
					previsualizaciones.setAll(pdfFileOut);
				} catch (OfficeException e) {
					crearAlerta(AlertType.WARNING, "Advertencia",
							"No se ha podido crear la previsualización \n" + "del documento importado.", null, false);
				}
				pdfCreado.set(pdfFileOut);
			}
		});

		// pageNumTextField - TextField del índice de registros/documentos creados
		// Añadiendo filtro numérico al pageNumTextField
		UnaryOperator<Change> filter = change -> {
			String text = change.getText();
			if (text.matches("[0-9]*")) {
				return change;
			}
			return null;
		};
		TextFormatter<String> textFormatter = new TextFormatter<>(filter);
		pageNumTextField.setTextFormatter(textFormatter);

		// listener de la listproperty previsualizaciones
		previsualizaciones.addListener((o, ov, nv) -> {
			if (nv.size() > 0) {
				pdfViewer.load(previsualizaciones.get(0));
				documentIndex.set(1);
				pageNumTextField.setText("1");
			}
		});

		// Evento para cambiar el índice del documento seleccionado al presionar enter
		// en el textField para cambiar los registros
		pageNumTextField.setOnKeyReleased(event -> {
			if (event.getCode() == KeyCode.ENTER) {
				int registro = Integer.parseInt(pageNumTextField.getText());
				if (registro <= previsualizaciones.size()//
						&& registro > 0 && !previsualizaciones.isEmpty()) {
					documentIndex.set(registro); // documentIndex tiene un listener
				}
			}
		});

		// Listener para cambiar el índice del documento seleccionado
		pageNumTextField.focusedProperty().addListener((o, ov, nv) -> {
			if (!nv) {
				try {
					int registro = Integer.parseInt(pageNumTextField.getText());
					if (registro <= previsualizaciones.size()//
							&& registro > 0 && !previsualizaciones.isEmpty()) {
						documentIndex.set(registro); // documentIndex tiene un listener
					}
				} catch (Exception e) {
				}
			}
		});

		// Listener para cuando se cambie el índice del documento seleccionado
		documentIndex.addListener((o, ov, nv) -> {
			if (nv != ov && nv.intValue() != 0) {
				try {
					pdfViewer.load(previsualizaciones.get(nv.intValue() - 1));
				} catch (Exception e) {
				}
			}
		});

		// CheckBoxs
		distinctMayusCheckBox.selectedProperty().addListener((o, ov, nv) -> {
			caseSensitive.set(nv);
		});
		exactWordsCheckBox.selectedProperty().addListener((o, ov, nv) -> {
			replaceExactWord.set(nv);
		});

		// bindings
		BooleanBinding rightButtonsBinding = Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (documentIndex.get() >= previsualizaciones.size() || previsualizaciones.size() == 0) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones);

		BooleanBinding leftButtonsBinding = Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (documentIndex.get() == 1 // || documentIndex.get() > previsualizaciones.size()
					|| previsualizaciones.size() == 0) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones);

		BooleanBinding previewButtonsBinding = Bindings.createBooleanBinding(() -> {
			if (previsualizaciones.size() == 0) {
				return true;
			} else {
				return false;
			}
		}, previsualizaciones);

		rightButton.disableProperty().bind(rightButtonsBinding);
		topRightButton.disableProperty().bind(rightButtonsBinding);
		leftButton.disableProperty().bind(leftButtonsBinding);
		topLeftButton.disableProperty().bind(leftButtonsBinding);

		pageNumTextField.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (previsualizaciones.size() == 0 || ficheroImportado.get() == null) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones, ficheroImportado));

		unloadButton.disableProperty().bind(previewButtonsBinding);
		reloadButton.disableProperty().bind(previewButtonsBinding);

		// Añadir la interfaz personalizada en Español al pdfViewer
		pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));

		// Crear setOnCloseRequest para cerrar los servicios de office y advertir de
		// salida
		WordSeedExporterApp.primaryStage.setOnCloseRequest(e -> {
			drawerController.salirApp();
			e.consume(); // Si llega hasta aquí es porque el usuario ha decidido cancelar la salida de la
							// aplicación
		});
		exactWordsCheckBox.requestFocus();
	}

	public Controller() {
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/View.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
		}
	}

	/**
	 * Cambia el tamaño de {@link JFXDrawer drawerMenu} a uno inferior para que se
	 * puedan seleccionar o interactuar con los elementos que hacen debajo del
	 * elemento
	 * 
	 * @param event El {@link JFXDrawerEvent} al que escucha
	 */
	@FXML
	void onDrawerClosed(JFXDrawerEvent event) {
		drawerMenu.setPrefWidth(280);
	}

	@FXML
	void onDrawerOpened(JFXDrawerEvent event) {
	}

	/**
	 * Cambia el documento previsualizado según el documento actual que está siendo
	 * previsualizado: Resta 1 a la propiedad documentIndex, lo que activa su
	 * listener el cuál cambia de documento según el valor de la propiedad.
	 * 
	 * También cambia el número mostrado en el cuadro de texto {@link TextArea
	 * pageNumTextField} a uno cuyo valor es su valor actual menos uno.
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onLeftButton(ActionEvent event) {
		if (documentIndex.get() > 1) {
			documentIndex.set(documentIndex.get() - 1);
			pageNumTextField.setText(documentIndex.get() + "");
		}
	}

	/**
	 * Cambia el documento previsualizado según el documento actual que está siendo
	 * previsualizado: Suma 1 a la propiedad documentIndex, lo que activa su
	 * listener el cuál cambia de documento según el valor de la propiedad.
	 * 
	 * También cambia el número mostrado en el cuadro de texto {@link TextArea
	 * pageNumTextField} a uno cuyo valor es su valor actual más uno.
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onRightButton(ActionEvent event) {
		if (documentIndex.get() < previsualizaciones.size() + 1) {
			documentIndex.set(documentIndex.get() + 1);
			pageNumTextField.setText(documentIndex.get() + "");
		}
	}

	/**
	 * Cambia el documento previsualizado a el primer documento de la lista de los
	 * documentos a previsualizar: Se establece a 1 la propiedad documentIndex, lo
	 * que activa su listener el cuál cambia de documento según el valor de la
	 * propiedad.
	 * 
	 * También cambia el número mostrado en el cuadro de texto {@link TextArea
	 * pageNumTextField} a uno cuyo valor es 1.
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onTopRightButton(ActionEvent event) {
		documentIndex.set(previsualizaciones.size());
		pageNumTextField.setText(documentIndex.get() + "");
	}

	/**
	 * Cambia el documento previsualizado a el último documento de la lista de los
	 * documentos a previsualizar: Se establece a el tamaño de esta lista la
	 * propiedad documentIndex, lo que activa su listener el cuál cambia de
	 * documento según el valor de la propiedad.
	 * 
	 * También cambia el número mostrado en el cuadro de texto {@link TextArea
	 * pageNumTextField} a uno cuyo valor es el tamaño de la lista mencionada.
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onTopLeftButton(ActionEvent event) {
		documentIndex.set(1);
		pageNumTextField.setText(documentIndex.get() + "");
	}

	/**
	 * Elimina la previsualización del documento que se está previsualizando
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onUnloadButton(ActionEvent event) {
		pdfViewer.unload();
	}

	/**
	 * Recarga la previsualización del documento que se está previsualizando
	 * 
	 * @param event El {@link ActionEvent} al que escucha
	 */
	@FXML
	void onReloadButton(ActionEvent event) {
		try {
			pdfViewer.load(previsualizaciones.get(documentIndex.get() - 1));
		} catch (Exception e) {
		}
	}

	/**
	 * Retorna la vista del elemento DrawerController
	 * 
	 * @return AnchorPane la vista del Conrolador Principal
	 */
	public AnchorPane getView() {
		return view;
	}

	/**
	 * Inicia el officeManager recibido para poder empezar posteriormente a usar los
	 * servicios de Office para la conversión de documentos a pdf y las
	 * previsualizaciones.
	 * 
	 * @param officeManager El {@link LocalOfficeManager} a recibir
	 */
	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
		try {
			this.officeManager.start();
			// Se crea un pdf vacío para luego usarlo y forzar a los servicios de Office a
			// iniciarse
			// Creando objeto documento-PDF
			PDDocument document = new PDDocument();
			// Añadiendo una página vacía
			document.addPage(new PDPage());
			try {
				String rutaFile1 = Controller.TEMPDOCSFOLDER.getPath() + File.separator + "ini.pdf";
				String rutaFile2 = Controller.TEMPDOCSFOLDER.getPath() + File.separator + "init.pdf";
				// Guardado
				document.save(rutaFile2);
				document.close();
				// Aquí se inician los servicios de Office en un nuevo hilo. No tiene utilidad
				// convertir un fichero pdf a pdf, solo sirve para forzar el inicio de los
				// servicios de libreoffice
				File f1 = new File(rutaFile1);
				File f2 = new File(rutaFile2);
				loadingBackground.setViewOrder(-1); // cambiar la pantalla de carga de servicios al plano frontal
				drawerMenu.setViewOrder(-2);
				new Thread(() -> {
					try {
						JodConverter.convert(f2)//
								.to(f1)//
								.as(DefaultDocumentFormatRegistry.PDF)//
								.execute();
					} catch (OfficeException | NullPointerException e) {
					}
					Platform.runLater(() -> { // se oculta la pantalla de carga
						loadingBackground.setViewOrder(0);
						drawerMenu.setViewOrder(0);
						try { // se eliminan los ficheros de inicio creados
							Files.delete(Paths.get(rutaFile1));
							Files.delete(Paths.get(rutaFile2));
						} catch (IOException e) {
						}
						previewReady.set(true);
					});
				}).start();
			} catch (IOException e) {
			}
			drawerController.setOfficeManager(this.officeManager);
		} catch (OfficeException e) {
			crearAlerta(AlertType.ERROR, "Error",
					"Se ha producido un error al iniciar los servicios de LibeOffice/OpenOffice", null, false);
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
	public static void crearAlerta(AlertType at, String t, String ht, String ct, boolean wait) {
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
	public static Optional<ButtonType> crearAlerta(Alert al, String t, String ht, String ct, boolean wait) {
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
}
