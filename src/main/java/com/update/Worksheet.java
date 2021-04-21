package com.update;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Worksheet {
	String path;
	
	
	public Worksheet(String path) {
		super();
		this.path = path;
	}

	
	public Map<String, String> getValuesFromWorksheet(){
		Map<String, String> hm= new LinkedHashMap<String, String>();
		
        try {
            FileInputStream file = new FileInputStream(new File(path));
  
            // Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
  
            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
  
            boolean flag=true;
            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
         	   
         	   
         	   if(flag==true) {
         		   rowIterator.next();
         		   flag=false;
         	   }
         	   
                Row row = rowIterator.next();
                // For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                	
                	
                while (cellIterator.hasNext()) {
                   hm.put(cellIterator.next().toString(),cellIterator.next().toString());
                }
            }
            workbook.close();
            file.close();
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        
        
        
        return hm;
    }     
}
