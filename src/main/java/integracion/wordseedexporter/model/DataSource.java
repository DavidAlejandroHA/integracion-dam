package integracion.wordseedexporter.model;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * Esta clase es un objeto que representa una tabla, en donde se almacenan por
 * separado las palabras clave que se quieren remplazar en el documento
 * importado, y las filas en donde se almacenan los distintos valores que se
 * quieren generar en el reemplazo de palabras.
 * 
 * Una vez que el lector de documentos .xlsx o .odt finaliza la lectura,
 * organizará la fuente de datos de forma que cada fila tendrá el mismo número
 * de elementos (Strings) que la lista de palabras clave, por lo que en caso de
 * faltar elementos en una de las filas se rellena automáticamente con strings
 * vacíos, los cuales no serán procesados por las funciones que utilizen las
 * fuentes de datos, en caso de haber palabras clave vacías, se elimina esta y
 * sus correspondientes índices de cada fila de la fuente de datos.
 */
public class DataSource {

	private ObservableList<ObservableList<String>> rows = FXCollections.observableArrayList();
	private ObservableList<String> keyNames = FXCollections.observableArrayList();

	public DataSource() {

	}

	public ObservableList<ObservableList<String>> getRows() {
		return rows;
	}

	public ObservableList<String> getKeyNames() {
		return keyNames;
	}

	public void addRow(ObservableList<String> obL) {
		rows.add(obL);
	}

	public void addKeyName(String s) {
		keyNames.add(s);
	}

	public void setRows(ObservableList<ObservableList<String>> rowList) {
		rows.setAll(rowList);
	}

	public void setKeyNames(ObservableList<String> keyNameList) {
		keyNames.setAll(keyNameList);
	}
}
