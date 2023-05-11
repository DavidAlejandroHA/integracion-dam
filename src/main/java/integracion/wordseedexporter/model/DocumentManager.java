package integracion.wordseedexporter.model;

import java.io.File;

import org.odftoolkit.odfdom.doc.OdfTextDocument;
import org.odftoolkit.odfdom.incubator.search.TextNavigation;
import org.odftoolkit.odfdom.incubator.search.TextSelection;

public class DocumentManager {
	
	public DocumentManager() {
		
	}
	
	public void giveDocument(File f) {
		/*try {
			System.out.println("a");
			System.out.println(FileMagic.valueOf(f));
			System.out.println("b");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println("sassdds");
		}*/
		try {
			//https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/OdfTextDocument.html
			//https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/package-summary.html // Tipos de documentos
			OdfTextDocument textdoc=(OdfTextDocument) OdfTextDocument.loadDocument(f);
			
			TextNavigation search1;

			search1= new TextNavigation("por",textdoc);

			while (search1.hasNext()) {

				TextSelection item1 = (TextSelection) search1.getCurrentItem();

				System.out.println(item1);

			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//https://docx4java.org/docx4j/Docx4j_GettingStarted.pdf
		//https://stackoverflow.com/questions/72405729/how-to-programatically-modify-libre-office-odt-document-in-java
		//https://angelozerr.wordpress.com/2012/12/06/how-to-convert-docxodt-to-pdfhtml-with-java/
		//https://stackoverflow.com/questions/61805246/docx4j-docx-to-pdf-conversion-docx-content-not-appearing-page-by-page-to-pdf
		
	}

}
