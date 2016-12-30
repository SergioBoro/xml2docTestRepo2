package ru.curs.xml2document;

import java.io.InputStream;
import java.util.Deque;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.transform.TransformerFactory;
import javax.xml.transform.sax.SAXResult;
import javax.xml.transform.stream.StreamSource;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

abstract class XMLDataReader {

	private static final Pattern XQUERY = Pattern
			.compile("([^\\[]+)\\[@([^=]+)=('([^']+)'|\"([^\"]+)\")\\]");

	private final ReportWriter writer;
	private final DescriptorElement descriptor;
	public static int countIteration = 0;

	XMLDataReader(ReportWriter writer, DescriptorElement descriptor) {
		this.writer = writer;
		this.descriptor = descriptor;
	}

	private enum ParserState {
		ELEMENT, ITERATION, OUTPUT
	}

	private static final class DescriptorParser extends DefaultHandler {

		private final Deque<DescriptorElement> elementsStack = new LinkedList<>();
		private DescriptorElement root;
		private ParserState parserState = ParserState.ITERATION;

		@Override
		public void startElement(String uri, String localName, String name,
				final Attributes atts) throws SAXException {

			abstract class AttrReader<T> {
				T getValue(String qName) throws XML2WordError {
					String buf = atts.getValue(qName);
					if (buf == null || "".equals(buf))
						return getIfEmpty();
					else
						return getIfNotEmpty(buf);
				}

				abstract T getIfNotEmpty(String value)
						throws XML2WordError;

				abstract T getIfEmpty();
			}

			final class StringAttrReader extends AttrReader<String> {
				@Override
				String getIfNotEmpty(String value) throws XML2WordError {
					return value;
				}

				@Override
				String getIfEmpty() {
					return null;
				}
			}

			try {
				switch (parserState) {
				case ELEMENT:
					if ("iteration".equals(localName)) {
						int index = (new AttrReader<Integer>() {
							@Override
							Integer getIfNotEmpty(String value) {
								return Integer.parseInt(value);
							}

							@Override
							Integer getIfEmpty() {
								return -1;
							}
						}).getValue("index");
						
						DescriptorIteration currIteration = new DescriptorIteration(
								index);
						elementsStack.peek().getSubelements()
								.add(currIteration);
						parserState = ParserState.ITERATION;
					} else if ("output".equals(localName)) {
						IDRange range = (new AttrReader<IDRange>() {
							@Override
							IDRange getIfNotEmpty(String value)
									throws XML2WordError {
								return new IDRange(value);
							}

							@Override
							IDRange getIfEmpty() {
								return null;
							}
						}).getValue("templateId");

						IDRanges ranges = (new AttrReader<IDRanges>() {
							@Override
							IDRanges getIfNotEmpty(String value)
									throws XML2WordError {
								return new IDRanges(value);
							}

							@Override
							IDRanges getIfEmpty() {
								return null;
							}
						}).getValue("templates");
						
						boolean pagebreak = (new AttrReader<Boolean>() {
							@Override
							Boolean getIfNotEmpty(String value)
									throws XML2WordError {
								return "true".equalsIgnoreCase(value);
							}

							@Override
							Boolean getIfEmpty() {
								return false;
							}
						}).getValue("pagebreak");
						
						//StringAttrReader sar = new StringAttrReader();
						
						DescriptorOutput output = new DescriptorOutput(range, ranges, pagebreak);
						elementsStack.peek().getSubelements().add(output);

						parserState = ParserState.OUTPUT;
					}
					break;
				case ITERATION:
					if ("element".equals(localName)) {
						String elementName = (new StringAttrReader())
								.getValue("name");
						DescriptorElement currElement = new DescriptorElement(
								elementName);

						if (root == null)
							root = currElement;
						else {
							// Добываем контекст текущей итерации...
							List<DescriptorSubelement> subelements = elementsStack
									.peek().getSubelements();
							DescriptorIteration iter = (DescriptorIteration) subelements
									.get(subelements.size() - 1);
							iter.getElements().add(currElement);
						}
						elementsStack.push(currElement);
						parserState = ParserState.ELEMENT;
					}
					break;
				}
			} catch (XML2WordError e) {
				throw new SAXException(e.getMessage());
			}
		}

		@Override
		public void endElement(String uri, String localName, String name)
				throws SAXException {
			switch (parserState) {
			case ELEMENT:
				elementsStack.pop();
				parserState = ParserState.ITERATION;
				break;
			case ITERATION:
				parserState = ParserState.ELEMENT;
				break;
			case OUTPUT:
				parserState = ParserState.ELEMENT;
				break;
			}
		}

	}

	static XMLDataReader createReader(InputStream xmlData,
			InputStream xmlDescriptor, boolean useSAX, ReportWriter writer)
			throws XML2WordError {
		if (xmlData == null)
			throw new XML2WordError("XML Data is null.");
		if (xmlDescriptor == null)
			throw new XML2WordError("XML descriptor is null.");

		DescriptorParser parser = new DescriptorParser();
		try {
			TransformerFactory
					.newInstance()
					.newTransformer()
					.transform(new StreamSource(xmlDescriptor),
							new SAXResult(parser));
		} catch (Exception e) {
			throw new XML2WordError(
					"Error while processing XML descriptor: " + e.getMessage());
		}
		if (useSAX)
			return new SAXDataReader(xmlData, parser.root, writer);
		else
			return new DOMDataReader(xmlData, parser.root, writer);
	}

	abstract void process() throws XML2WordError;

	final void processOutput(XMLContext c, DescriptorOutput o)
			throws XML2WordError {
		if (o.getRange() != null || o.getRanges() != null)	
			getWriter().section(c, o.getRange(), o.getRanges(), o.getPageBreak());
	}

	static final boolean compareIndices(int expected, int actual) {
		return (expected < 0) || (actual == expected);
	}

	static final boolean compareNames(String expected, String actual,
			Map<String, String> attributes) {
		if (expected == null)
			return actual == null;

		if (expected.startsWith("*") || expected.equals(actual))
			return true;

		Matcher m = XQUERY.matcher(expected);
		if (m.matches()) {
			String tagName = m.group(1);
			if (!tagName.equals(actual))
				return false;
			String attribute = m.group(2);
			String value = m.group(4) == null ? m.group(5) : m.group(4);
			return value.equals(attributes.get(attribute));
		} else
			return false;
	}

	final ReportWriter getWriter() {
		return writer;
	}

	final DescriptorElement getDescriptor() {
		return descriptor;
	}

	static final class DescriptorElement {
		private final String elementName;
		private final List<DescriptorSubelement> subelements = new LinkedList<DescriptorSubelement>();

		public DescriptorElement(String elementName) {
			this.elementName = elementName;
		}

		String getElementName() {
			return elementName;
		}

		List<DescriptorSubelement> getSubelements() {
			return subelements;
		}
	}

	abstract static class DescriptorSubelement {
	}

	static final class DescriptorIteration extends DescriptorSubelement {
		private final int index;
		private final List<DescriptorElement> elements = new LinkedList<>();

		public DescriptorIteration(int index) {
			this.index = index;
		}

		int getIndex() {
			return index;
		}

		List<DescriptorElement> getElements() {
			return elements;
		}
	}

	static final class DescriptorOutput extends DescriptorSubelement {
		private final IDRange range;
		private final IDRanges ranges;
		private final boolean pageBreak;
		
		public DescriptorOutput(
				IDRange range, IDRanges ranges, boolean pageBreak) throws XML2WordError {
			this.range = range;
			this.ranges = ranges;
			this.pageBreak = pageBreak;
			}

		public IDRange getRange() {
			return range;
		}
		
		public IDRanges getRanges() {
			return ranges;
		}
		
		public boolean getPageBreak() {
			return pageBreak;
		}
	}

}
