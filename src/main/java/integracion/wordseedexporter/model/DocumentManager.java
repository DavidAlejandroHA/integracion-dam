package integracion.wordseedexporter.model;

import java.io.File;
import java.io.IOException;

import org.apache.poi.poifs.filesystem.FileMagic;

public class DocumentManager {
	
	public DocumentManager() {
		
	}
	
	public void giveDocument(File f) {
		try {
			System.out.println("a");
			System.out.println(FileMagic.valueOf(f));
			System.out.println("b");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println("sassdds");
		}
		
	}

}
