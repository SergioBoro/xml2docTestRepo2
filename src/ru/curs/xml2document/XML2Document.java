package ru.curs.xml2document;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;


public class XML2Document {

	private XML2Document() {
	}
	
	public static void process(String xmlData, String xmlDescriptor,
			String temlpate, String output) throws FileNotFoundException, IOException, XML2WordError
	{
		ReportWriter writer = ReportWriter.createWriter(new FileInputStream(temlpate), OutputType.DOCX,
				false, new FileOutputStream(output));
		XMLDataReader reader = XMLDataReader.createReader(new FileInputStream(xmlData),
				new FileInputStream(xmlDescriptor), false, writer);
		reader.process();
	}
	
	public static void process(InputStream xmlData, InputStream xmlDescriptor,
			InputStream template, OutputStream output) throws FileNotFoundException, IOException, XML2WordError
	{
		ReportWriter writer = ReportWriter.createWriter(template, OutputType.DOCX,
				false, output);
		XMLDataReader reader = XMLDataReader.createReader(xmlData,
				xmlDescriptor, false, writer);
		reader.process();
	}
	
	
}
