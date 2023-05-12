package integracion.wordseedexporter.model;

import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class ReplaceWordHandler extends DefaultHandler {

	private String key1;
	private String key2;
	private StringBuilder content;

	public ReplaceWordHandler(String k1, String k2){
		key1 = k1;
		key2 = k2;
		content = new StringBuilder();
	}
	
	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		String contentS = new String(ch, start, length);
		content.append(contentS);
		/*if(content.contains(key1)) {
			content = content.replaceAll(key1, key2);
			ch = content.toCharArray();
			//System.out.println(ch.toString());
		}*/
		/*if(!contentS.equals(content)) {
			System.out.println("a");
		}*/
		//System.out.println(content);
		//TODO: Usar DOM 
		super.characters(ch, start, length);
	}
	
	//https://stackoverflow.com/questions/41832382/sax-parser-returns-empty-string-from-xml
}
