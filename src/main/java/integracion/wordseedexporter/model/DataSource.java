package integracion.wordseedexporter.model;

import javafx.collections.ObservableList;

public class DataSource {

	private ObservableList<ObservableList<String>> rows;
	private ObservableList<String> keyNames;

	public DataSource() {

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
