package ru.curs.xml2document;

import java.util.ArrayList;
import java.util.List;


final public class IDRanges {
	private String idsString;
	
	IDRanges(String idsString)
	{
		this.idsString = idsString;
	}
	
	public List<Integer> getIDs()
	{
		List<Integer> list = new ArrayList<>();
		String[] array = this.idsString.split(":");
		for(String str : array)
		{
			list.add(Integer.parseInt(str));
		}
		return list;
	}
	
}
