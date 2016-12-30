package ru.curs.xml2document;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

abstract class ReportWriter {
	
	private OutputStream output;
	
	static ReportWriter createWriter(InputStream template, OutputType type,
			boolean copyTemplate, OutputStream output)
			throws XML2WordError {
		InputStream localTemplate = template;
		InputStream templateCopy = null;
		try {
			if (copyTemplate) {
				XML2WordBLOB b = new XML2WordBLOB(template);
				localTemplate = b.getInStream();
				templateCopy = b.getInStream();
				template.close();
			}
		} catch (IOException e) {
			throw new XML2WordError(e.getMessage());
		}
		ReportWriter result = null;
		switch (type) {
		case DOC:
			//result = DOCReportWriter(localTemplate, templateCopy);
			//Not yet implemented
			break;
		case DOCX:
			result = new DOCXReportWriter(localTemplate, templateCopy);
			break;
		default:
			// This will never happen
			return null;
		}
		result.output = output;
		return result;
	}
	
	public void section(XMLContext context,
			IDRange range, IDRanges ranges, boolean pagebreak) throws XML2WordError {
		if(range != null)
			putSection(context, range);
		
		if(ranges != null)
		{
			for(int id = ranges.getIDs().get(0); id <= ranges.getIDs().get(1); id++)
				putSections(context, id);
		}
	}

	private void processPageBreak(boolean pagebreak) {
		//Not yet implemented
	}
	
	final OutputStream getOutput() {
		return output;
	}
	
	abstract void putSection(XMLContext context, IDRange range) throws XML2WordError;

	abstract void putSections(XMLContext context, int id) throws XML2WordError;
	
	public abstract void flush() throws XML2WordError;

}
