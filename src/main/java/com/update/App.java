package com.update;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class App 
{
    public static void main( String[] args ) throws FileNotFoundException, InvalidFormatException, IOException
    {
    	//path of document file
        String pathOfDoc="G:\\MavenProjects\\UpdateIt\\OldFiles\\oldOne.docx";
        
        //path of excel sheet
        String pathOfValues="G:\\MavenProjects\\UpdateIt\\Excel Sheets\\data.xlsx";
        
        //Class for fetching data from excel sheet
        Worksheet worksheet=new Worksheet(pathOfValues);
        //above Class object will return a hashmap;
        Map<String, String> hm=worksheet.getValuesFromWorksheet();
      
//        System.out.println(hm);
        //Class for modification of document file on the basis of excel sheet data
        Replace replace=new Replace(pathOfDoc,hm);
//        after modification it will return the path of updated file
        	String newPath=replace.replaceIt();
        	System.out.println("Path of updated doc:- "+newPath);
    }
}
