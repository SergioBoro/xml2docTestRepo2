package ru.curs.xml2document.test;

import java.io.FileNotFoundException;
import java.io.IOException;
import ru.curs.xml2document.XML2WordError;
import ru.curs.xml2document.XML2Document;

public final class Main {

	public static void main(String[] args) throws FileNotFoundException, IOException, XML2WordError {
		XML2Document.process("filesForTesting/dataXML.xml", "filesForTesting/descriptorXML.xml",
				"filesForTesting/template.docx", "filesForTesting/result.docx");
	}

}
