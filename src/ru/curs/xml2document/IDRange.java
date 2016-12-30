package ru.curs.xml2document;


final public class IDRange {
	private String idString;
	
	IDRange(String idString)
	{
		this.idString = idString;
	}
	
	public int getID()
	{
		return Integer.parseInt(this.idString);
	}
	
}
