package ru.curs.xml2document;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;

import java.io.*;
import java.math.BigInteger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;


public class DOCXReportWriter extends ReportWriter {

	private XWPFDocument template = null;
	private CustomXWPFDocument result = null;
	final Map<Integer, XWPFRun> runMap = new HashMap<>();
	List<Integer> intList = new ArrayList<>();
	Map<Integer, Integer> beginTemplateMap = new HashMap<>();
	Map<Integer, Integer> endTemplateMap = new HashMap<>();
	Map<Integer, CTText> ctMap = new HashMap<>();
	static int m = -1;
	HashMap<Pair, String> cellMap = new HashMap<>();
	Map<HashMap<Integer, Integer>, String> intParMap = new HashMap<>();
	Map<HashMap<Integer, Integer>, String> intStrMap = new HashMap<>();
	
	List<HashMap<Integer, Integer>> placeholderPosList = new ArrayList<>();
	List<HashMap<String, String>> textList = new ArrayList<>();
	String flagText = "";
	List<String> strL = new ArrayList<>();
	Boolean f = false;
	String font = "";
	
	public DOCXReportWriter(InputStream template, InputStream templateCopy)
			throws XML2WordError
	{
		try {
			this.template = new XWPFDocument(template);
			this.result = new CustomXWPFDocument();
			
			if(this.template.getDocument().getBody() != null
					&& this.template.getDocument().getBody().getSectPr() != null
					&& this.template.getDocument().getBody().getSectPr().getType() != null)
			{	
				this.result.getDocument().getBody().addNewSectPr().addNewType().setVal(
						this.template.getDocument().getBody().getSectPr().getType().getVal());
			}
			
			try {
				this.result.createStyles();
				(this.result.getStyles()).setStyles(this.template.getStyle());
			} catch (XmlException e) {
				throw new XML2WordError(e.getMessage());
			}
			
			copyLayout(this.template, this.result);
	        
			XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(result,
					result.getDocument().getBody().addNewSectPr());
			if(this.template.getHeaderFooterPolicy() != null && 
					this.template.getHeaderFooterPolicy().getDefaultFooter() != null)
			{
				headerFooterPolicy.createFooter(STHdrFtr.DEFAULT).
					setHeaderFooter(this.template.getHeaderFooterPolicy().getDefaultFooter()._getHdrFtr());
				
			}
			else if(this.template.getHeaderFooterPolicy() != null && this.template.getFooterList() != null
					&& this.template.getFooterList().size() > 0)
			{
				XWPFParagraph[] prg = this.template.getFooterList().get(0).getParagraphs().
						toArray(new XWPFParagraph[0]);
				
				headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, prg);
				//headerFooterPolicy.createFooter(STHdrFtr.DEFAULT).
					//setHeaderFooter(this.template.getFooterList().get(0)._getHdrFtr());
			}
//			else if(this.template.getHeaderFooterPolicy() != null &&
//					this.template.getHeaderFooterPolicy().getFirstPageFooter() != null)
//			{
//				headerFooterPolicy.createFooter(STHdrFtr.FIRST).
//				setHeaderFooter(this.template.getHeaderFooterPolicy().getFirstPageFooter()._getHdrFtr());
//			}
				
			if(this.template.getHeaderFooterPolicy() != null && 
					this.template.getHeaderFooterPolicy().getDefaultHeader() != null)
			{
				headerFooterPolicy.createHeader(STHdrFtr.DEFAULT).
					setHeaderFooter(this.template.getHeaderFooterPolicy().getDefaultHeader()._getHdrFtr());
				
			}
			else if(this.template.getHeaderFooterPolicy() != null && this.template.getHeaderList() != null
					&& this.template.getHeaderList().size() > 0)
			{
				XWPFParagraph[] prg = this.template.getHeaderList().get(0).getParagraphs().
						toArray(new XWPFParagraph[0]);
				
				headerFooterPolicy.createHeader(STHdrFtr.DEFAULT, prg);
				//headerFooterPolicy.createHeader(STHdrFtr.DEFAULT).
					//setHeaderFooter(this.template.getHeaderList().get(0)._getHdrFtr());
			}
//			else if(this.template.getHeaderFooterPolicy() != null && 
//					this.template.getHeaderFooterPolicy().getFirstPageHeader() != null)
//			{
//				headerFooterPolicy.createHeader(STHdrFtr.FIRST).
//					setHeaderFooter(this.template.getHeaderFooterPolicy().getFirstPageHeader()._getHdrFtr());
//				
//			}
			
			for(XWPFFootnote footNote : this.template.getFootnotes())
			{
				this.result.createFootnotes();
				this.result.addFootnote(footNote.getCTFtnEdn());
			}
			
		} catch (IOException e) {
			throw new XML2WordError(e.getMessage());
		}
	}
	
	@Override
	void putSection(XMLContext context, IDRange range) throws XML2WordError 
	{
		for (int i = 0; i < this.template.getBodyElements().size(); i++) {
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH 
			 && ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("<template"))
			{
				String str = ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().
						replaceAll(">[0-9]+", "").replaceAll(".*<template", "").replaceAll("[^0-9]", "");
				m = Integer.parseInt(str);
				beginTemplateMap.put(m, i);
			}
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH
					&& ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("</template"))
			{
				endTemplateMap.put(m, i);
			}
		}
		
		cycle: for (int i = 0; i < this.template.getBodyElements().size(); i++) {
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH 
			 && !((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("<template"))
			{
				for(XWPFRun rn : ((XWPFParagraph)this.template.getBodyElements().get(i)).getRuns())
            	{	
            		if(rn != null && rn.getCTR().getRPr() != null) 
            		{
            			if(rn.getCTR().getRPr().getRFonts() != null
    						&& rn.getCTR().getRPr().getRFonts().getAscii() != null)
    					{
    						font = rn.getCTR().getRPr().getRFonts().getAscii();
    					}
            	
    					else if(rn.getCTR().getRPr().getRFonts() != null
        						&& rn.getCTR().getRPr().getRFonts().getEastAsia() != null)
            			{
            				font = rn.getCTR().getRPr().getRFonts().getEastAsia();
            			}
            			
    					else if(rn.getCTR().getRPr().getRFonts() != null
        						&& rn.getCTR().getRPr().getRFonts().getHAnsi() != null)
            			{
            				font = rn.getCTR().getRPr().getRFonts().getHAnsi();
            			}
            		}
            		
            		if(rn.getFontFamily() != null)
            		{
            			font = rn.getFontFamily();
            		}
            		
            		if(!("".equals(font)) && font != null)
            			break cycle;
            	}
			}
		}
		
		if("".equals(font) || font == null || "SimSun".equals(font))
			font = "Times New Roman";
		
		int j = range.getID();
		for (int g = beginTemplateMap.get(j); g <= endTemplateMap.get(j); g++)
		{
			IBodyElement bodyElement = template.getBodyElements().get(g); 

            BodyElementType elementType = bodyElement.getElementType();

            if (elementType == BodyElementType.PARAGRAPH) {
            	
                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;
                
                for(XWPFRun r : srcPr.getRuns())
                {
                	if(r instanceof XWPFHyperlinkRun)
                		throw new XML2WordError("Document contains hyperlinks");
                }
                
                if(srcPr.getParagraphText().contains("<template") &&
            			srcPr.getParagraphText().contains("paragraph"))
            	{ 
                	continue;
                }
            			
            	if(srcPr.getParagraphText().contains("<template") &&
                		srcPr.getParagraphText().contains("table")) 
            	{
            		continue;
            	}
            				
                if(srcPr.getParagraphText().contains("</template") &&
                        !(srcPr.getParagraphText().contains("letters"))) 
                { 
                	continue;
                }
                
                 
                 if(srcPr.getParagraphText().contains("<template") && srcPr.getParagraphText().contains("letters"))
                 {
                	 int preBeginInd = srcPr.getParagraphText().indexOf(">");
                	 int preEndInd = srcPr.getParagraphText().indexOf("</");
                	 String letText = srcPr.getParagraphText().substring(preBeginInd + 1, preEndInd);
                	 strL.add(letText);
                	 continue;
                 }
                 
                 if(!srcPr.getParagraphText().contains("<template") && strL.size()>0)
                 {
                	 String res = "";
                	 for(String s : strL)
                		 res +=s;
                		 
                	 srcPr.insertNewRun(0).setText(res, 0);
                	 strL.clear();
                	 f = true;
                 }
                  
                copyStyle(template, result, template.getStyles().getStyle(srcPr.getStyleID()));
                  
                boolean hasImage = false;

                XWPFParagraph dstPr = result.createParagraph();
                
                for (XWPFRun srcRun : srcPr.getRuns()) {

                    dstPr.createRun();

                    if (srcRun.getEmbeddedPictures().size() > 0)
                        hasImage = true;

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                    	byte[] img = pic.getPictureData().getData();

                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
                        
                        try {
                           String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),
                                    Document.PICTURE_TYPE_PNG);
                            result.createPictureCxCy(srcPr, blipId, result.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
                                    cx, cy);
                            
                        } catch (InvalidFormatException e1) {
                            e1.printStackTrace();
                        }
                    }
                }

                if (hasImage == false)
                {
                	
                	while(srcPr.getParagraphText().contains("~{")){
                		
                		HashMap<Integer, Integer> placeholderPosMap = new HashMap<>();
                		HashMap<String, String> textMap = new HashMap<>();
                		
                		TextSegement tsStart = srcPr.searchText("~{", new PositionInParagraph());
                		TextSegement tsEnd = srcPr.searchText("}", new PositionInParagraph());
                		int startRunPos = tsStart.getBeginRun();
                		int endRunPos = tsEnd.getEndRun();
                	
                		placeholderPosMap.put(startRunPos, endRunPos);
                		placeholderPosList.add(placeholderPosMap);
                    
                		StringBuilder sb = new StringBuilder();
                    
                		for(int q = startRunPos; q <= endRunPos; q++)
                		{
                			XWPFRun r = srcPr.getRuns().get(q);
                			int pos = r.getTextPosition();
                			if(r.getText(pos) != null) {
                				sb.append(r.getText(pos));
                			}
                		}
                	
                		if(sb.toString().contains("@") && !(sb.toString().contains("/")))
                		{
                			int index1 = sb.toString().indexOf("~{");
                			int index2 = sb.toString().indexOf("}");
                			String xPath = sb.toString().substring(index1, index2 + 1);
                			int index = xPath.indexOf("@");
                			if(!((xPath.substring(index-1, index)).equals("{")))
                			{
                				throw new XML2WordError(xPath + " contains incorrect x-path structure");
                			}
                		}
                		
                		String runText = context.calc(sb.toString());
                		flagText = runText;
                	
                		textMap.put(runText, sb.toString());
                		textList.add(textMap);
                	
                		srcPr.getRuns().get(startRunPos).setText(runText, 0);
                		srcPr.getRuns().get(startRunPos).setFontFamily(font);
                		for(int q = endRunPos; q > startRunPos; q--)
                		{
                			srcPr.getRuns().get(q).setText("", 0);
                		}
                	}
                
                	for(XWPFRun rn : srcPr.getRuns())
                 	{	
             				rn.setFontFamily(font);
                 	}
                	
                	dstPr.setStyle(srcPr.getStyle());
                    
                	int pos = result.getParagraphs().size() - 1;
                    	result.setParagraph(srcPr, pos);
                    
                    if(f){
                    	srcPr.removeRun(0);
                    	f = false;
                    }
                    
                    if(srcPr.getParagraphText().contains(flagText)){
                    	it: for(int e = 0; e < placeholderPosList.size(); e++)
                    	{
                    		try
                    		{
                    			HashMap<Integer, Integer> plHM= placeholderPosList.get(e);
                    			HashMap<String, String> tHM = textList.get(e);
                    	
                    			int startRunPos = (Integer)plHM.keySet().toArray()[0];
                    			int endRunPos = plHM.get(startRunPos);
                    	
                    			String calcText = (String)tHM.keySet().toArray()[0];
                    			String initText = tHM.get(calcText);
                    	
                    			srcPr.getRuns().get(startRunPos).setText(initText, 0);
                    			for(int q = endRunPos; q > startRunPos; q--)
                    			{
                    				srcPr.getRuns().get(q).setText("", 0);
                    			}
                    		}
                    		catch(Exception ex)
                    		{
                    			break it;
                    		}
                    	}
                    }
                }

                textList.clear();
                placeholderPosList.clear();
                
            } else if (elementType == BodyElementType.TABLE) {
            	
                XWPFTable table = (XWPFTable) bodyElement;
                
                for(XWPFTableRow row : table.getRows())
                {
                	for(XWPFTableCell cell : row.getTableCells())
                	{
                		for(XWPFParagraph par : cell.getParagraphs())
                		{
                			for(XWPFRun r : par.getRuns())
                            {
                            	if(r instanceof XWPFHyperlinkRun)
                            		throw new XML2WordError("Document contains hyperlinks");
                            }
                		}
                	}
                }

                XWPFTable dstTbl = result.createTable();
                
                copyStyle(template, result, template.getStyles().getStyle(table.getStyleID()));
                
                dstTbl.setStyleID(table.getStyleID());
				
                String initText = "";
                
                for(int d=0; d<table.getRows().size(); d++)
                {
                	XWPFTableRow row = table.getRows().get(d);
                	for(int m = 0; m<row.getTableCells().size(); m++)
                	{
                		XWPFTableCell cell = row.getTableCells().get(m);
                		String cellText = cell.getText();
                		//for(int z = 0; z < cell.getParagraphs().size(); z++){
                			if(cellText.contains("~{"))
                			{
                				initText = cellText;
                				cellMap.put(new Pair(d, m), initText);
                			}
                		//}
                	}
                }
                
				if(table.getText().contains("~{"))
				{
					for(Pair x : cellMap.keySet())
					{
							XWPFTableRow row2 = table.getRows().get(x.getFirst());
							XWPFTableCell cell = row2.getTableCells().get(x.getSecond());
							String color = cell.getColor();
							for(int z = 0; z < cell.getParagraphs().size(); z++)
							{
								XWPFParagraph par = cell.getParagraphs().get(z);
                       	  		if(par.getParagraphText().contains("~{"))
                       	  		{
                       	  			HashMap<Integer, Integer> hashM = new HashMap<Integer, Integer>();
                       	  			hashM.put(x.getSecond(), z);
                       	  			List<XWPFRun> runs = par.getRuns();
                       	  			intParMap.put(hashM, par.getParagraphText());
                       	  			String calc = context.calc(par.getParagraphText());
                       	  			intStrMap.put(hashM, calc);
                       	  			runs.get(0).setText(calc, 0);
                       	  			for(int i = 1; i < runs.size(); i++)
                       	  			{
                       	  				runs.get(i).setText("", 0);
                       	  			}
                       	  		}
							}
							cell.setColor(color);
					}	
               }
                
				for(XWPFTableRow tRow : table.getRows())
				{
					for(XWPFTableCell cell : tRow.getTableCells())
					{
						for(XWPFParagraph par : cell.getParagraphs())
						{
							for(XWPFRun rn : par.getRuns())
							{
								if(!("".equals(font)) && font != null)
									rn.setFontFamily(font);
							}
						}
					}
				}
				
                int pos = result.getTables().size() - 1;
                result.setTable(pos, table);
              
                for(int d=0; d<table.getRows().size(); d++)
                {
                	XWPFTableRow row = table.getRows().get(d);
                	for(int m = 0; m<row.getTableCells().size(); m++)
                	{
                		XWPFTableCell cell = row.getTableCells().get(m);
                		for(int z = 0; z < cell.getParagraphs().size(); z++)
						{
							XWPFParagraph par = cell.getParagraphs().get(z);
                   	  		for(HashMap<Integer, Integer> c : intStrMap.keySet())
                   	  		{
                   	  			if(c.keySet().contains(m) && par.getParagraphText().equals(intStrMap.get(c)))
                   	  			{
                   	  				List<XWPFRun> runs = par.getRuns();	
                   	  				if(runs.size() > 0)
                   	  				{
                   	  					runs.get(0).setText(intParMap.get(c), 0);
                   	  					for(int i = 1; i < runs.size(); i++)
                   	  					{
                   	  						runs.get(i).setText("", 0);
                   	  					}
                   	  				}
                   	  			}
                   	  		}
						}
                	}
                }
                
                
//                for(Pair x : cellMap.keySet())
//				{		
//                	
//					XWPFTableRow row2 = table.getRows().get(x.getFirst());
//					XWPFTableCell cell = row2.getTableCells().get(x.getSecond());
//					String color = cell.getColor();
//					//if(cell.getText().contains(cellFlagText))
//					//{
//						for(int z = 0; z < cell.getParagraphs().size(); z++)
//						{
//							XWPFParagraph par = cell.getParagraphs().get(z);
//                   	  		List<XWPFRun> runs = par.getRuns();
//                   	  		if(runs.size() > 0)
//                   	  		{
//                   	  			runs.get(0).setText(cellMap.get(x), 0);
//                   	  			for(int i = 0; i < runs.size(); i++)
//                   	  			{
//                   	  				runs.get(i).setText("", 0);
//                   	  			}
//                   	  		}
//						}
//						//cell.setText(cellMap.get(x));
//					//}
//						cell.setColor(color);
//                	
//				}
//                
                cellMap.clear();
                intParMap.clear();
                intStrMap.clear();
            }
            
   		}
	}

	@Override
	void putSections(XMLContext context, int id) throws XML2WordError 
	{
		for (int i = 0; i < this.template.getBodyElements().size(); i++) {
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH 
			 && ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("<template"))
			{
				String str = ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().
						replaceAll(">[0-9]+", "").replaceAll(".*<template", "").replaceAll("[^0-9]", "");
				m = Integer.parseInt(str);
				beginTemplateMap.put(m, i);
			}
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH
					&& ((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("</template"))
			{
				endTemplateMap.put(m, i);
			}
		}
		
		cycle: for (int i = 0; i < this.template.getBodyElements().size(); i++) {
			if(this.template.getBodyElements().get(i).getElementType() == BodyElementType.PARAGRAPH 
			 && !((XWPFParagraph)this.template.getBodyElements().get(i)).getText().contains("<template"))
			{
				for(XWPFRun rn : ((XWPFParagraph)this.template.getBodyElements().get(i)).getRuns())
            	{	
            		if(rn != null && rn.getCTR().getRPr() != null) 
            		{
            			if(rn.getCTR().getRPr().getRFonts() != null
    						&& rn.getCTR().getRPr().getRFonts().getAscii() != null)
    					{
    						font = rn.getCTR().getRPr().getRFonts().getAscii();
    					}
            	
    					else if(rn.getCTR().getRPr().getRFonts() != null
        						&& rn.getCTR().getRPr().getRFonts().getEastAsia() != null)
            			{
            				font = rn.getCTR().getRPr().getRFonts().getEastAsia();
            			}
            			
    					else if(rn.getCTR().getRPr().getRFonts() != null
        						&& rn.getCTR().getRPr().getRFonts().getHAnsi() != null)
            			{
            				font = rn.getCTR().getRPr().getRFonts().getHAnsi();
            			}
            		}
            		
            		if(rn.getFontFamily() != null)
            		{
            			font = rn.getFontFamily();
            		}
            		
            		if(!("".equals(font)) && font != null)
            			break cycle;
            	}
			}
		}
		
		if("".equals(font) || font == null || "SimSun".equals(font))
			font = "Times New Roman";
		
		int j = id;
		for (int g = beginTemplateMap.get(j); g <= endTemplateMap.get(j); g++)
		{
			IBodyElement bodyElement = template.getBodyElements().get(g); 

            BodyElementType elementType = bodyElement.getElementType();

            if (elementType == BodyElementType.PARAGRAPH) {
            	
                XWPFParagraph srcPr = (XWPFParagraph) bodyElement;
                
                for(XWPFRun r : srcPr.getRuns())
                {
                	if(r instanceof XWPFHyperlinkRun)
                		throw new XML2WordError("Document contains hyperlinks");
                }
                
                if(srcPr.getParagraphText().contains("<template") &&
            			srcPr.getParagraphText().contains("paragraph"))
            	{ 
                	continue;
                }
            			
            	if(srcPr.getParagraphText().contains("<template") &&
                		srcPr.getParagraphText().contains("table")) 
            	{
            		continue;
            	}
            				
                if(srcPr.getParagraphText().contains("</template") &&
                        !(srcPr.getParagraphText().contains("letters"))) 
                { 
                	continue;
                }
                
                 
                 if(srcPr.getParagraphText().contains("<template") && srcPr.getParagraphText().contains("letters"))
                 {
                	 int preBeginInd = srcPr.getParagraphText().indexOf(">");
                	 int preEndInd = srcPr.getParagraphText().indexOf("</");
                	 String letText = srcPr.getParagraphText().substring(preBeginInd + 1, preEndInd);
                	 strL.add(letText);
                	 continue;
                 }
                 
                 if(!srcPr.getParagraphText().contains("<template") && strL.size()>0)
                 {
                	 String res = "";
                	 for(String s : strL)
                		 res +=s;
                		 
                	 srcPr.insertNewRun(0).setText(res, 0);
                	 strL.clear();
                	 f = true;
                 }
                 
                copyStyle(template, result, template.getStyles().getStyle(srcPr.getStyleID()));
                  
                boolean hasImage = false;

                XWPFParagraph dstPr = result.createParagraph();
                
                for (XWPFRun srcRun : srcPr.getRuns()) {

                    dstPr.createRun();

                    if (srcRun.getEmbeddedPictures().size() > 0)
                        hasImage = true;

                    for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                    	byte[] img = pic.getPictureData().getData();

                        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
                        
                        try {
                           String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),
                                    Document.PICTURE_TYPE_PNG);
                            result.createPictureCxCy(srcPr, blipId, result.getNextPicNameNumber(Document.PICTURE_TYPE_PNG),
                                    cx, cy);
                            
                        } catch (InvalidFormatException e1) {
                            e1.printStackTrace();
                        }
                    }
                }

                if (hasImage == false)
                {
                	
                	while(srcPr.getParagraphText().contains("~{")){
                		
                		HashMap<Integer, Integer> placeholderPosMap = new HashMap<>();
                		HashMap<String, String> textMap = new HashMap<>();
                		
                		TextSegement tsStart = srcPr.searchText("~{", new PositionInParagraph());
                		TextSegement tsEnd = srcPr.searchText("}", new PositionInParagraph());
                		int startRunPos = tsStart.getBeginRun();
                		int endRunPos = tsEnd.getEndRun();
                	
                		placeholderPosMap.put(startRunPos, endRunPos);
                		placeholderPosList.add(placeholderPosMap);
                    
                		StringBuilder sb = new StringBuilder();
                    
                		for(int q = startRunPos; q <= endRunPos; q++)
                		{
                			XWPFRun r = srcPr.getRuns().get(q);
                			int pos = r.getTextPosition();
                			if(r.getText(pos) != null) {
                				sb.append(r.getText(pos));
                			}
                		}
                	
                		if(sb.toString().contains("@") && !(sb.toString().contains("/")))
                		{
                			int index1 = sb.toString().indexOf("~{");
                			int index2 = sb.toString().indexOf("}");
                			String xPath = sb.toString().substring(index1, index2 + 1);
                			int index = xPath.indexOf("@");
                			if(!((xPath.substring(index-1, index)).equals("{")))
                			{
                				throw new XML2WordError(xPath + " contains incorrect x-path structure");
                			}
                		}
                		
                		String runText = context.calc(sb.toString());
                		flagText = runText;
                	
                		textMap.put(runText, sb.toString());
                		textList.add(textMap);
                	
                		srcPr.getRuns().get(startRunPos).setText(runText, 0);
                		srcPr.getRuns().get(startRunPos).setFontFamily(font);
                		for(int q = endRunPos; q > startRunPos; q--)
                		{
                			srcPr.getRuns().get(q).setText("", 0);
                		}
                	}
                
                	for(XWPFRun rn : srcPr.getRuns())
                 	{	
             				rn.setFontFamily(font);
                 	}
                	
                	dstPr.setStyle(srcPr.getStyle());
                    
                	int pos = result.getParagraphs().size() - 1;
                    	result.setParagraph(srcPr, pos);
                    
                    if(f){
                    	srcPr.removeRun(0);
                    	f = false;
                    }
                    
                    if(srcPr.getParagraphText().contains(flagText)){
                    	it: for(int e = 0; e < placeholderPosList.size(); e++)
                    	{
                    		try
                    		{
                    			HashMap<Integer, Integer> plHM= placeholderPosList.get(e);
                    			HashMap<String, String> tHM = textList.get(e);
                    	
                    			int startRunPos = (Integer)plHM.keySet().toArray()[0];
                    			int endRunPos = plHM.get(startRunPos);
                    	
                    			String calcText = (String)tHM.keySet().toArray()[0];
                    			String initText = tHM.get(calcText);
                    	
                    			srcPr.getRuns().get(startRunPos).setText(initText, 0);
                    			for(int q = endRunPos; q > startRunPos; q--)
                    			{
                    				srcPr.getRuns().get(q).setText("", 0);
                    			}
                    		}
                    		catch(Exception ex)
                    		{
                    			break it;
                    		}
                    	}
                    }
                }

                textList.clear();
                placeholderPosList.clear();
                
            } else if (elementType == BodyElementType.TABLE) {
            	
                XWPFTable table = (XWPFTable) bodyElement;

                for(XWPFTableRow row : table.getRows())
                {
                	for(XWPFTableCell cell : row.getTableCells())
                	{
                		for(XWPFParagraph par : cell.getParagraphs())
                		{
                			for(XWPFRun r : par.getRuns())
                            {
                            	if(r instanceof XWPFHyperlinkRun)
                            		throw new XML2WordError("Document contains hyperlinks");
                            }
                		}
                	}
                }
                
                XWPFTable dstTbl = result.createTable();
                
                copyStyle(template, result, template.getStyles().getStyle(table.getStyleID()));
                
                dstTbl.setStyleID(table.getStyleID());
				
                String initText = "";
                
                for(int d=0; d<table.getRows().size(); d++)
                {
                	XWPFTableRow row = table.getRows().get(d);
                	for(int m = 0; m<row.getTableCells().size(); m++)
                	{
                		XWPFTableCell cell = row.getTableCells().get(m);
                		String cellText = cell.getText();
                			if(cellText.contains("~{"))
                			{
                				initText = cellText;
                				cellMap.put(new Pair(d, m), initText);
                			}
                	}
                }
                
				if(table.getText().contains("~{"))
				{
					for(Pair x : cellMap.keySet())
					{
							XWPFTableRow row2 = table.getRows().get(x.getFirst());
							XWPFTableCell cell = row2.getTableCells().get(x.getSecond());
							String color = cell.getColor();
							for(int z = 0; z < cell.getParagraphs().size(); z++)
							{
								XWPFParagraph par = cell.getParagraphs().get(z);
                       	  		if(par.getParagraphText().contains("~{"))
                       	  		{
                       	  			HashMap<Integer, Integer> hashM = new HashMap<Integer, Integer>();
                       	  			hashM.put(x.getSecond(), z);
                       	  			List<XWPFRun> runs = par.getRuns();
                       	  			intParMap.put(hashM, par.getParagraphText());
                       	  			String calc = context.calc(par.getParagraphText());
                       	  			intStrMap.put(hashM, calc);
                       	  			runs.get(0).setText(calc, 0);
                       	  			for(int i = 1; i < runs.size(); i++)
                       	  			{
                       	  				runs.get(i).setText("", 0);
                       	  			}
                       	  		}
							}
							cell.setColor(color);
					}	
               }
                
				for(XWPFTableRow tRow : table.getRows())
				{
					for(XWPFTableCell cell : tRow.getTableCells())
					{
						for(XWPFParagraph par : cell.getParagraphs())
						{
							for(XWPFRun rn : par.getRuns())
							{
								if(!("".equals(font)) && font != null)
									rn.setFontFamily(font);
							}
						}
					}
				}
				
                int pos = result.getTables().size() - 1;
                result.setTable(pos, table);
              
                for(int d=0; d<table.getRows().size(); d++)
                {
                	XWPFTableRow row = table.getRows().get(d);
                	for(int m = 0; m<row.getTableCells().size(); m++)
                	{
                		XWPFTableCell cell = row.getTableCells().get(m);
                		for(int z = 0; z < cell.getParagraphs().size(); z++)
						{
							XWPFParagraph par = cell.getParagraphs().get(z);
                   	  		for(HashMap<Integer, Integer> c : intStrMap.keySet())
                   	  		{
                   	  			if(c.keySet().contains(m) && par.getParagraphText().equals(intStrMap.get(c)))
                   	  			{
                   	  				List<XWPFRun> runs = par.getRuns();	
                   	  				if(runs.size() > 0)
                   	  				{
                   	  					runs.get(0).setText(intParMap.get(c), 0);
                   	  					for(int i = 1; i < runs.size(); i++)
                   	  					{
                   	  						runs.get(i).setText("", 0);
                   	  					}
                   	  				}
                   	  			}
                   	  		}
						}
                	}
                }

                cellMap.clear();
                intParMap.clear();
                intStrMap.clear();
            }
            
   		}
	}

	
	 private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style)
	    {
	        if (destDoc == null || style == null)
	            return;

	        if (destDoc.getStyles() == null) {
	            destDoc.createStyles();
	        }

	        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
	        for (XWPFStyle xwpfStyle : usedStyleList) {
	            destDoc.getStyles().addStyle(xwpfStyle);
	        }
	    }

	 private static void copyLayout(XWPFDocument srcDoc, XWPFDocument destDoc)
	    {
	        CTPageMar pgMar = srcDoc.getDocument().getBody().getSectPr().getPgMar();

	        BigInteger bottom = pgMar.getBottom();
	        BigInteger footer = pgMar.getFooter();
	        BigInteger gutter = pgMar.getGutter();
	        BigInteger header = pgMar.getHeader();
	        BigInteger left = pgMar.getLeft();
	        BigInteger right = pgMar.getRight();
	        BigInteger top = pgMar.getTop();

	        CTPageMar addNewPgMar = destDoc.getDocument().getBody().addNewSectPr().addNewPgMar();

	        addNewPgMar.setBottom(bottom);
	        addNewPgMar.setFooter(footer);
	        addNewPgMar.setGutter(gutter);
	        addNewPgMar.setHeader(header);
	        addNewPgMar.setLeft(left);
	        addNewPgMar.setRight(right);
	        addNewPgMar.setTop(top);

	        CTPageSz pgSzSrc = srcDoc.getDocument().getBody().getSectPr().getPgSz();

	        BigInteger code = pgSzSrc.getCode();
	        BigInteger h = pgSzSrc.getH();
	        org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation.Enum orient = pgSzSrc.getOrient();
	        BigInteger w = pgSzSrc.getW();

	        CTPageSz addNewPgSz = destDoc.getDocument().getBody().addNewSectPr().addNewPgSz();

	        addNewPgSz.setCode(code);
	        addNewPgSz.setH(h);
	        addNewPgSz.setOrient(orient);
	        addNewPgSz.setW(w);
	    }
	 
	@Override
	public void flush() throws XML2WordError {
		try {
			result.write(getOutput());
		} catch (IOException e) {
			throw new XML2WordError(e.getMessage());
		}
	}

	XWPFDocument getResult() {
		return result;
	}
	
	public class Pair
	{
		private int first;
		private int second;
		
		Pair(int aFirst, int aSecond)
		{
			first = aFirst;
			second = aSecond;
		}
		
		public int getFirst()
		{
			return first;
		}
		
		public int getSecond()
		{
			return second;
		}
		
		public void setFirst(int aFirst)
		{
			first = aFirst;
		}
		
		public void setSecond(int aSecond)
		{
			second = aSecond;
		}
	}
	
	public class CustomXWPFDocument extends XWPFDocument
	{
	    public CustomXWPFDocument() throws IOException
	    {
	        super();
	    }

	    public void createPictureCxCy(XWPFParagraph srcPr, String blipId,int id, long cx, long cy)
	    {
	    	XWPFParagraph ppp = createParagraph();
	    	ppp.setAlignment(srcPr.getAlignment());
	    	ppp.setSpacingAfter(srcPr.getSpacingAfter());
	    	CTInline inline = ppp.createRun().getCTR().addNewDrawing().addNewInline();

	        String picXml = "" +
	                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
	                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
	                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
	                "         <pic:nvPicPr>" +
	                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
	                "            <pic:cNvPicPr/>" +
	                "         </pic:nvPicPr>" +
	                "         <pic:blipFill>" +
	                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
	                "            <a:stretch>" +
	                "               <a:fillRect/>" +
	                "            </a:stretch>" +
	                "         </pic:blipFill>" +
	                "         <pic:spPr>" +
	                "            <a:xfrm>" +
	                "               <a:off x=\"0\" y=\"0\"/>" +
	                "               <a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>" +
	                "            </a:xfrm>" +
	                "            <a:prstGeom prst=\"rect\">" +
	                "               <a:avLst/>" +
	                "            </a:prstGeom>" +
	                "         </pic:spPr>" +
	                "      </pic:pic>" +
	                "   </a:graphicData>" +
	                "</a:graphic>";

	        XmlToken xmlToken = null;
	        try
	        {
	            xmlToken = XmlToken.Factory.parse(picXml);
	        }
	        catch(XmlException xe)
	        {
	            xe.printStackTrace();
	        }
	        inline.set(xmlToken);

	        inline.setDistT(0);
	        inline.setDistB(0);
	        inline.setDistL(0);
	        inline.setDistR(0);
	        
	        CTPositiveSize2D extent = inline.addNewExtent();
	        extent.setCx(cx);
	        extent.setCy(cy);

	        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
	        docPr.setId(id);
	        docPr.setName("Picture " + id);
	        docPr.setDescr("Generated");
	    }

	    public void createPicture(XWPFParagraph srcPr, String blipId,int id, int width, int height)
	    {
	        final int EMU = 9525;
	        width *= EMU;
	        height *= EMU;
	     
	        createPictureCxCy(srcPr, blipId, id, width, height);
	    }
	}
}


