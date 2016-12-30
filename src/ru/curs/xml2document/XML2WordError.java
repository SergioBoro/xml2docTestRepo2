package ru.curs.xml2document;

public class XML2WordError  extends Exception {
	private static final long serialVersionUID = 4382588062277186741L;

	public XML2WordError(String string) {
		super("В процессе обработки документа произошла следующая ошибка: " + string);
	}

}
