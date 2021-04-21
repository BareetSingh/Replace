package com.update;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Replace {
	String path;
	Map<String, String> hm= new HashMap<String, String>();
	
	public Replace(String path,Map<String, String> values) {
		super();
		this.path = path;
		this.hm = values;
	}
	

	public String replaceIt() throws FileNotFoundException, IOException, InvalidFormatException {
		
		XWPFDocument doc = new XWPFDocument(OPCPackage.open(path));
		
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            
		            
		            for (Map.Entry<String, String> me : hm.entrySet()) {
		            	if (text != null && text.contains(me.getKey())) {
			                text = text.replace(me.getKey(),me.getValue());
			                r.setText(text, 0);
			            }
		            }
		            
		            
		        }
		    }
		}
		
		for (XWPFTable tbl : doc.getTables()) {
		   for (XWPFTableRow row : tbl.getRows()) {
		      for (XWPFTableCell cell : row.getTableCells()) {
		         for (XWPFParagraph p : cell.getParagraphs()) {
		            for (XWPFRun r : p.getRuns()) {
		              String text = r.getText(0);
		              for (Map.Entry<String, String> me : hm.entrySet()) {
		              
		              if (text != null && text.contains(me.getKey())) {
		                text = text.replace(me.getKey(), me.getValue());
		                r.setText(text,0);
		              }
		              }
		            }
		         }
		      }
		   }
		}
		
		
		String newPath="G:\\MavenProjects\\UpdateIt\\UpdatedFiles\\updatedOne.docx";
		doc.write(new FileOutputStream(newPath));
		doc.close();
		return newPath;

	}

}	
