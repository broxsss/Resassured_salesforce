package com.amazon.restapi.Restassured;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExelUtil {
	
	XSSFWorkbook wb = null;
	
	public XSSFWorkbook getExcelFile(File file){
		
		try {
			 wb = new XSSFWorkbook(new FileInputStream(file));
			return wb;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return wb;
	}
	
	public List<Map<String, String>> getExcelList(XSSFWorkbook wb,String headerData){
		
	  List<Map<String,String>> tempList = new ArrayList<Map<String,String>>();
	  Map<Integer,String> header = new HashMap<Integer,String>();
	  String[] arg =headerData.split(",");
	
	  for(int i=0;i<arg.length; i++){
		 if(!arg[i].equals(""))
		   header.put(i, arg[i]);
	  }
	  
	  for( Row row : wb.getSheetAt(0) ) {
		
		  Map<String,String>  tempMap = new HashMap<String,String>();
		   for(Cell cell :row){
			   if(header.get(cell.getColumnIndex()) != null){
		   		   switch( cell.getCellType()) {
	                case Cell.CELL_TYPE_STRING :
	                    tempMap.put( header.get(cell.getColumnIndex()),cell.getRichStringCellValue().getString());
	                    break;
	                case Cell.CELL_TYPE_NUMERIC :
	                    if(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell))
	                    tempMap.put(  header.get(cell.getColumnIndex()),cell.getDateCellValue().toString());
	                    else
	                    	tempMap.put(  header.get(cell.getColumnIndex()),Integer.toString((int)cell.getNumericCellValue()));
	                    break;
	                case Cell.CELL_TYPE_FORMULA :
	                		tempMap.put(  header.get(cell.getColumnIndex()),cell.getCellFormula());
	                		break;
		   		   }
			   }
		   }
		   
		   tempList.add(tempMap);
	  }
	  
	  return tempList;
	}

}