package ru.curs.xml2document;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public final class XML2WordBLOB {
	private DataPage data;
	private boolean isModified;
	private int size;

	public XML2WordBLOB() {
	}

	XML2WordBLOB(final InputStream source) throws IOException {
		InputStream counter = new InputStream() {
			@Override
			public int read() throws IOException {
				int result = source.read();
				if (result >= 0)
					size++;
				return result;
			}
		};
		int buf = counter.read();
		data = buf < 0 ? new DataPage(0) : new DataPage(buf, counter);
	}

	public XML2WordBLOB clone() {
		XML2WordBLOB result = new XML2WordBLOB();
		result.data = data;
		result.size = size;
		return result;
	}

	public boolean isModified() {
		return isModified;
	}

	public InputStream getInStream() {
		return data == null ? null : data.getInStream();
	}

	public OutputStream getOutStream() {
		isModified = true;
		data = new DataPage();
		size = 0;
		return new OutputStream() {
			private DataPage tail = data;

			@Override
			public void write(int b) {
				tail = tail.write(b);
				size++;
			}
		};
	}

	public boolean isNull() {
		return data == null;
	}

	public void setNull() {
		isModified = isModified || (data != null);
		size = 0;
		data = null;
	}

	public int size() {
		return size;
	}

	private static final class DataPage {
		private static final int DEFAULT_PAGE_SIZE = 0xFFFF;
		private static final int BYTE_MASK = 0xFF;

		private final byte[] data;
		private DataPage nextPage;
		private int pos;

		DataPage() {
			this(DEFAULT_PAGE_SIZE);
		}

		private DataPage(int size) {
			data = new byte[size];
		}

		private DataPage(int firstByte, InputStream source) throws IOException {
			this();
			int buf = firstByte;
			while (pos < data.length && buf >= 0) {
				data[pos++] = (byte) buf;
				buf = source.read();
			}
			nextPage = buf < 0 ? null : new DataPage(buf, source);
		}

		DataPage write(int b) {
			if (pos < data.length) {
				data[pos++] = (byte) (b & BYTE_MASK);
				return this;
			} else {
				DataPage result = new DataPage();
				nextPage = result;
				return result.write(b);
			}
		}

		InputStream getInStream() {
			return new InputStream() {
				private int i = 0;
				private DataPage currentPage = DataPage.this;

				@Override
				public int read() {
					if (i < currentPage.pos)
						return (int) currentPage.data[i++] & BYTE_MASK;
					else if (currentPage.nextPage != null) {
						i = 0;
						currentPage = currentPage.nextPage;
						return read();
					} else {
						return -1;
					}
				}
			};
		}
	}
}