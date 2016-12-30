package ru.curs.xml2document;

import java.io.InputStream;
import java.util.HashMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

final class DOMDataReader extends XMLDataReader {

	private final Document xmlData;

	DOMDataReader(InputStream xmlData, DescriptorElement xmlDescriptor,
			ReportWriter writer) throws XML2WordError {
		super(writer, xmlDescriptor);
		try {
			DocumentBuilder db = DocumentBuilderFactory.newInstance()
					.newDocumentBuilder();
			this.xmlData = db.parse(xmlData);
		} catch (Exception e) {
			throw new XML2WordError("Error while parsing input data: "
					+ e.getMessage());
		}
	}

	private void processElement(String elementPath, DescriptorElement de,
			Element xe, int position) throws XML2WordError {
		XMLContext context = null;
		for (DescriptorSubelement se : de.getSubelements()) {
			if (se instanceof DescriptorIteration) {
				processIteration(elementPath, xe, (DescriptorIteration) se,
						position);
				countIteration++;
			} else if (se instanceof DescriptorOutput) {
				if (context == null)
					context = new XMLContext.DOMContext(xe, elementPath,
							position);
				processOutput(context, (DescriptorOutput) se);
			}
		}
	}

	private void processIteration(String elementPath, Element parent,
			DescriptorIteration i, int position) throws XML2WordError {

		final HashMap<String, Integer> elementIndices = new HashMap<>();

		for (DescriptorElement de : i.getElements())
			if ("(before)".equals(de.getElementName()))
				processElement(elementPath, de, parent, position);

		Node n = parent.getFirstChild();
		int elementIndex = -1;

		int pos = 0;
		iteration: while (n != null) {
			if (n.getNodeType() == Node.ELEMENT_NODE) {
				Integer ind = elementIndices.get(n.getNodeName());
				if (ind == null)
					ind = 0;
				elementIndices.put(n.getNodeName(), ind + 1);

				elementIndex++;
				boolean found = false;
				if (compareIndices(i.getIndex(), elementIndex)) {
					HashMap<String, String> atts = new HashMap<>();

					for (int j = 0; j < n.getAttributes().getLength(); j++) {
						Node att = n.getAttributes().item(j);
						atts.put(att.getNodeName(), att.getNodeValue());
					}

					for (DescriptorElement e : i.getElements())
						if (compareNames(e.getElementName(), n.getNodeName(),
								atts)) {
							found = true;
							processElement(String.format("%s/%s[%s]",
									elementPath, n.getNodeName(),
									elementIndices.get(n.getNodeName())
											.toString()), e, (Element) n,
									pos + 1);
						}
					if (i.getIndex() >= 0)
						break iteration;
				}
				if (found)
					pos++;
			}
			n = n.getNextSibling();
		}

		for (DescriptorElement de : i.getElements())
			if ("(after)".equals(de.getElementName()))
				processElement(elementPath, de, parent, position);
	}

	@Override
	void process() throws XML2WordError {
		if (getDescriptor().getElementName().equals(
				xmlData.getDocumentElement().getNodeName()))
			processElement("/" + getDescriptor().getElementName() + "[1]",
					getDescriptor(), xmlData.getDocumentElement(), 1);
		getWriter().flush();
	}
}
