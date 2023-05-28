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

	public static ListProperty<DataSource> dataSources = new SimpleListProperty<>(FXCollections.observableArrayList());

	public static ObjectProperty<File> ficheroImportado = new SimpleObjectProperty<>();

	public static ListProperty<File> previsualizaciones = new SimpleListProperty<>(FXCollections.observableArrayList());

	private static ObjectProperty<File> pdfCreado = new SimpleObjectProperty<>();

	private static BooleanProperty previewReady = new SimpleBooleanProperty(false);

	private static IntegerProperty documentIndex = new SimpleIntegerProperty(1);

	public static boolean officeInstalled = true; // valor por defecto

	@Override
	public void initialize(URL location, ResourceBundle resources) {
		// load data
		drawerController = new DrawerController();
		drawerMenu.setSidePane(drawerController.getView());
		drawerMenu.setPrefWidth(280);
		drawerController.setDrawerMenu(drawerMenu);
		pdfViewer.setDisable(true);

		rightButton.setDisable(true);
		leftButton.setDisable(true);
		pageNumTextField.setDisable(true);
		pageNumTextField.setText("1");

		// crear carpeta de la app si no existe
		if (!Controller.APPFOLDER.exists()) {
			Controller.APPFOLDER.mkdirs();
			Controller.TEMPDOCSFOLDER.mkdirs();
		}

		// listeners
		pdfCreado.addListener((o, ov, nv) -> {
			// El if y el else es para forzar al pdfViewer que cambie de pdf
			if (nv != null) {
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
			} else {
				pdfViewer.unload();
				pdfViewer.setDisable(true);
			}
			HBox hb = (HBox) view.lookup("#pdfHBox");
			hb.getChildren().set(1, pdfViewer);
		});

		// listener para el fichero importado
		// en este caso se intenta hacer una previsualización del documento en el visor
		// de pdfs
		ficheroImportado.addListener((o, ov, nv) -> {
			if (nv != null && previewReady.get()) {
//				Dialog<Boolean> waitingDialog = new Dialog<>();
//				ProgressBar progressBar = new ProgressBar();
//				progressBar.setProgress(ProgressIndicator.INDETERMINATE_PROGRESS);
//				waitingDialog.getDialogPane().setContent(progressBar);
//				waitingDialog.getDialogPane().setPrefWidth(300);
//				waitingDialog.getDialogPane().setCenterShape(true);
//				waitingDialog.getDialogPane().getStylesheets()
//						.add(getClass().getResource("/css/darkstyle.css").toExternalForm());
//				waitingDialog.setTitle("Cargando...");
//				waitingDialog.setHeaderText("Iniciando servicios de Office");
//				waitingDialog.show();
//				try {
//					Files.delete(Paths.get(Controller.TEMPDOCSFOLDER + File.separator + "preview.pdf"));
//				} catch (IOException e1) {
//					e1.printStackTrace();
//				}
				File pdfFileOut = new File(Controller.TEMPDOCSFOLDER + File.separator + "preview.pdf");
				try {
					JodConverter.convert(nv)// .as(DefaultDocumentFormatRegistry.DOC)
							.to(pdfFileOut)//
							.as(DefaultDocumentFormatRegistry.PDF)//
							.execute();
				} catch (OfficeException e) {
					crearAlerta(AlertType.WARNING, "Advertencia",
							"No se ha podido crear la previsualización \n" + "del documento importado.", null, false);
				}
				// El null es para forzar al pdfViewer que cambie de pdf
				pdfCreado.set(null);
				pdfCreado.set(pdfFileOut);
				// waitingDialog.setResult(true);
				// waitingDialog.close();
				// });
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

		previsualizaciones.addListener((o, ov, nv) -> {
			if (nv.size() > 1) {
				pdfViewer.load(previsualizaciones.get(0));
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
			if (nv != ov) {
				pdfViewer.load(previsualizaciones.get(nv.intValue() - 1));
			}
		});

		// bindings
		rightButton.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (documentIndex.get() >= previsualizaciones.size() || previsualizaciones.size() == 0) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones));

		leftButton.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (documentIndex.get() == 1 || documentIndex.get() > previsualizaciones.size()
					|| previsualizaciones.size() == 0) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones));

		pageNumTextField.disableProperty().bind(Bindings.createBooleanBinding(() -> {
			boolean limitReached = false;
			if (previsualizaciones.size() == 0 || ficheroImportado.get() == null) {
				limitReached = true;
			}
			return limitReached;
		}, documentIndex, previsualizaciones, ficheroImportado));

		// Añadir la interfaz personalizada en Español al pdfViewer
		pdfViewer.setSkin(new PDFViewSkinES(pdfViewer));

		// Crear setOnCloseRequest para cerrar los servicios de office y advertir de
		// salida
		WordSeedExporterApp.primaryStage.setOnCloseRequest(e -> {
			drawerController.salirApp();
			e.consume(); // Si llega hasta aquí es porque el usuario ha decidido cancelar la salida de la
							// aplicación
		});
	}

	public Controller() {
		try {
			FXMLLoader loader = new FXMLLoader(getClass().getResource("/fxml/View.fxml"));
			loader.setController(this);
			loader.load();
		} catch (IOException e) {
		}
	}

	@FXML
	void onDrawerClosed(JFXDrawerEvent event) {
		drawerMenu.setPrefWidth(280);
	}

	@FXML
	void onDrawerOpened(JFXDrawerEvent event) {
		// drawerMenu.setPrefWidth(600);
	}

	@FXML
	void onLeftButton(ActionEvent event) {
		if (documentIndex.get() > 1) {
			documentIndex.set(documentIndex.get()-1);
			pageNumTextField.setText(documentIndex.get() + "");
		}
	}

	@FXML
	void onRightButton(ActionEvent event) {
		if (documentIndex.get() < previsualizaciones.size() + 1) {
			documentIndex.set(documentIndex.get()+1);;
			pageNumTextField.setText(documentIndex.get() + "");
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

	public void setOfficeManager(LocalOfficeManager officeManager) {
		this.officeManager = officeManager;
		try {
			this.officeManager.start();
//			crearAlerta(AlertType.INFORMATION, "Iniciando servicios de Office",
//					"Se están iniciando los servicios de Office",
//					"Una vez iniciados, se podrán previsualizar documentos\n" + "y exportar en formato pdf.", false);
			// Se crea un pdf vacío para luego usarlo y forzar a los servicios de Office a
			// iniciarse
			// Creando objeto documento-PDF
			PDDocument document = new PDDocument();

			// Añadiendo una página vacía
			document.addPage(new PDPage());

			// Guardado
			try {
				String rutaFile1 = Controller.TEMPDOCSFOLDER.getPath() + File.separator + "ini.pdf";
				String rutaFile2 = Controller.TEMPDOCSFOLDER.getPath() + File.separator + "init.pdf";
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

					// A veces en la primera previsualización no sale bien, por lo que se fuerza una
					// previsualización momentánea
					Platform.runLater(() -> {
						loadingBackground.setViewOrder(0);
						drawerMenu.setViewOrder(0);
						// pdfViewer.load(f2);
						// pdfViewer.unload();
						try {
							Files.delete(Paths.get(rutaFile2));
						} catch (IOException e) {
						}
						previewReady.set(true);
					});
				}).start();
			} catch (IOException e) {
//				Platform.runLater(() -> {
//					crearAlerta(AlertType.ERROR, "Error", "No se han podido cargar los servicios de Office.",
//							"Es posible que al importar un documento la próxima previsualización\n"
//									+ "no funcione correctamente y/o tarde en efectuarse\n"
//									+ "el inicio de los servicios.",
//							false);
//				});
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
