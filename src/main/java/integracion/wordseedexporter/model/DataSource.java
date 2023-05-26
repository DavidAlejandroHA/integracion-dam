package integracion.wordseedexporter.model;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

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
