package com.amazon.restapi.Restassured;



import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;
import org.apache.http.HttpStatus;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.opencsv.CSVWriter;

import  io.restassured.response.*;
import static io.restassured.RestAssured.given;
import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


public class OrderActivation_TimeDuration_account {

	private static XSSFWorkbook workbook;
	private static String accesstoken=null;
	private static utilityproperties ul = new utilityproperties();
	static TreeMap<String, ArrayList<String>> map = new TreeMap<String, ArrayList<String>>();
	static ArrayList<String> arrlist = new ArrayList<String>();

	public static void writeDataAtOnce(Map<String, ArrayList<String>> map2,String macdoutputfile) 
	{ 

		// first create file object for file placed at location  
		File file = new File(System.getProperty("user.dir")+"//"+macdoutputfile+".csv"); 
		try { 
			// create FileWriter object with file as parameter 
			FileWriter outputfile = new FileWriter(file); 

			// create CSVWriter object filewriter object as parameter 
			CSVWriter writer = new CSVWriter(outputfile, ',', 
					CSVWriter.NO_QUOTE_CHARACTER, 
					CSVWriter.DEFAULT_ESCAPE_CHARACTER, 
					CSVWriter.DEFAULT_LINE_END);

			for (Entry<String, ArrayList<String>> ee : map2.entrySet()) {
				List<String> values = ee.getValue();
				String str[] = new String[values.size()]; 

				// ArrayList to Array Conversion 
				for (int j = 0; j < values.size(); j++) { 

					// Assign each value to String array 
					str[j] = values.get(j); 
				}
				List<String[]> list = new ArrayList<String[]>();
				list.add(str);
				writer.writeAll(list); 
			} 
			writer.close(); 
		} 
		catch (IOException e) { 
			e.printStackTrace(); 
		} 
	}


	private static String authenticateUser() throws IOException {

		//Getting the details from properties file
		String url= ul.getData("tokenurl");
		String password=ul.getData("password");
		String username=ul.getData("username");
		String client_secret=ul.getData("client_secret");
		String client_id=ul.getData("client_id");


		//Getting the access Token
		String response =
				given()
				.header("Content-Type", "application/x-www-form-urlencoded").formParam("grant_type", "password")
				.formParam("client_id", client_id)
				.formParam("client_secret", client_secret)
				.formParam("username", username)
				.formParam("password", password)
				.when()
				.post(url)
				.asString();

		//System.out.println(response);
		JsonPath jsonPath = new JsonPath(response);
		String accessToken = jsonPath.getString("access_token");

		return accessToken;
	}

	//Time Difference between two date
	public static long getDateDiff(Date date1, Date date2, TimeUnit timeUnit) {
		long diffInMillies = date2.getTime() - date1.getTime();
		return timeUnit.convert(diffInMillies,TimeUnit.MILLISECONDS);
	}

	//Writing the Excel to CSV
	public static void echoAsCSV(Sheet sheet,FileWriter file) throws IOException {
		Row row = null;
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				//System.out.print(""+row.getCell(j)+",");
				String data = ""+row.getCell(j)+",";
				file.write(data);
			}
			file.write("\r\n");
			//System.out.println();
		}
	}

	//Reading the Excel to write in CSV file
	public static void exceltocsv(String sheetName,String workbookname){
		InputStream inp = null;
		try {
			inp = new FileInputStream(System.getProperty("user.dir")+"//"+workbookname+".xlsx");
			FileWriter file = new FileWriter(new File(System.getProperty("user.dir")+"//"+workbookname+"_"+sheetName+".csv"));
			Workbook wb = WorkbookFactory.create(inp);

			for(int i=0;i<wb.getNumberOfSheets();i++) {
				System.out.println(wb.getSheetAt(i).getSheetName());
				echoAsCSV(wb.getSheetAt(i),file);
			}
			file.close();
		} catch (InvalidFormatException ex) {
			Logger.getLogger(OrderActivation_TimeDuration_account.class.getName()).log(Level.SEVERE, null, ex);
			System.out.println("Not able to write in CSV");
		} catch (FileNotFoundException ex) {
			Logger.getLogger(OrderActivation_TimeDuration_account.class.getName()).log(Level.SEVERE, null, ex);
			System.out.println("Not able to write in CSV");
		} catch (IOException ex) {
			Logger.getLogger(OrderActivation_TimeDuration_account.class.getName()).log(Level.SEVERE, null, ex);
			System.out.println("Not able to write in CSV");
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				Logger.getLogger(OrderActivation_TimeDuration_account.class.getName()).log(Level.SEVERE, null, ex);
				System.out.println("Not able to write in CSV");
			}
		}
	}

	//Write to excel sheet with hashmap data object created in result method 
	public static void writeExcel(String sheetName,String workbookname,Map<Integer,Object[]> data) throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		//Check if workbook with same name exists
		if(new File(System.getProperty("user.dir")+"//"+workbookname+".xlsx").exists())
		{
			boolean flag = false ;
			FileInputStream inputStream = new FileInputStream(new File(System.getProperty("user.dir")+"//"+workbookname+".xlsx"));
			workbook = (XSSFWorkbook) WorkbookFactory.create(inputStream);

			XSSFCellStyle Dstyle = workbook.createCellStyle();
			Dstyle.setBorderBottom(BorderStyle.THIN);
			Dstyle.setBorderRight(BorderStyle.THIN);
			Dstyle.setBorderLeft(BorderStyle.THIN);

			XSSFCellStyle Hstyle = workbook.createCellStyle();
			Dstyle.setBorderTop(BorderStyle.THIN);
			Hstyle.setBorderBottom(BorderStyle.THIN);
			Hstyle.setBorderRight(BorderStyle.THIN);
			Hstyle.setBorderLeft(BorderStyle.THIN);
			Hstyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			Hstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			Hstyle.setLocked(true);

			Font Hfont = workbook.createFont();
			Hfont.setColor(IndexedColors.BLACK.getIndex());
			Hstyle.setFont(Hfont);

			Font Dfont = workbook.createFont();
			Dfont.setColor(IndexedColors.BLACK.getIndex());
			Dstyle.setFont(Dfont);
			// Create a blank sheet
			if (workbook.getNumberOfSheets() != 0) {
				for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
					flag = workbook.getSheetName(i).equals(sheetName);
				}
				if (flag) {
					XSSFSheet  sheet = workbook.getSheet(sheetName);
					// Iterate over data and write to sheet
					Set<Integer> keyset = data.keySet();
					int rownum = 0;
					for (Integer key : keyset) {
						// this creates a new row in the sheet
						if(key==0){
							Row row = sheet.getRow(rownum++);
							int cellnum = row.getLastCellNum();
							Object[] objArr = (Object[]) data.get(key);
							for(int i=3;i<objArr.length;i++){
								sheet.setColumnWidth(cellnum,5000);
								// this line append a cell in the next column of that row
								Cell cell = row.createCell(cellnum++);
								if (objArr[i] instanceof String){
									cell.setCellValue((String)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);
								}
								else if (objArr[i] instanceof Float){
									cell.setCellValue((Float)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);

								}
								else if (objArr[i] instanceof Integer){
									cell.setCellValue((Integer)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);
								}
							}
						}
						else{
							Row row = sheet.getRow(rownum++);
							int cellnum = row.getLastCellNum();
							sheet.setColumnWidth(cellnum,5000);
							Object[] objArr = (Object[]) data.get(key);
							for(int i=3;i<objArr.length;i++){
								// this line creates a cell in the next column of that row
								Cell cell = row.createCell(cellnum++);
								if (objArr[i] instanceof String){
									cell.setCellValue((String)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
								else if (objArr[i] instanceof Float){
									cell.setCellValue((Float)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
								else if (objArr[i] instanceof Integer){
									cell.setCellValue((Integer)objArr[i]);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
							}
						}
					}
					try {
						inputStream.close();
						//Write on the Excel with the given workspace name
						File file = new File(System.getProperty("user.dir")+"//"+workbookname+".xlsx");
						FileOutputStream out = new FileOutputStream(file);
						workbook.write(out);
						out.close();
					}
					catch (Exception e) {
						e.printStackTrace();
					}
				}
				else{
					XSSFSheet sheet = workbook.createSheet(sheetName);
					// Iterate over data and write to sheet
					Set<Integer> keyset = data.keySet();
					int rownum = 0;
					for (Integer key : keyset) {// this creates a new row in the sheet
						if(key==0){
							//Row row = sheet.getRow(rownum++);
							Row row = sheet.createRow(rownum++);
							int cellnum = 0;
							sheet.setColumnWidth(cellnum,5000);
							Object[] objArr = (Object[]) data.get(key);
							for (Object obj : objArr) {
								// this line creates a cell in the next column of that row
								Cell cell = row.createCell(cellnum++);
								if (obj instanceof String){
									cell.setCellValue((String)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);
								}
								else if (obj instanceof Float){
									cell.setCellValue((Float)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);
								}
								else if (obj instanceof Integer){
									cell.setCellValue((Integer)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Hstyle);
								}
							}
						}
						else{
							Row row = sheet.createRow(rownum++);
							int cellnum = 0;
							sheet.setColumnWidth(cellnum,5000);
							Object[] objArr = (Object[]) data.get(key);
							for (Object obj : objArr) {
								// this line creates a cell in the next column of that row
								Cell cell = row.createCell(cellnum++);
								if (obj instanceof String){
									cell.setCellValue((String)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
								else if (obj instanceof Float){
									cell.setCellValue((Float)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
								else if (obj instanceof Integer){
									cell.setCellValue((Integer)obj);
									sheet.setColumnWidth(cellnum,5000);
									cell.setCellStyle(Dstyle);
								}
							}
						}
					}
					try {
						inputStream.close();
						//Write on the Excel with the given workspace name
						File file = new File(System.getProperty("user.dir")+"//"+workbookname+".xlsx");
						FileOutputStream out = new FileOutputStream(file);
						workbook.write(out);
						out.close();

					}
					catch (Exception e) {
						e.printStackTrace();
					}
				}
			}
			else{
				// Create a blank sheet
				XSSFSheet sheet = workbook.createSheet(sheetName);

				// Iterate over data and write to sheet
				Set<Integer> keyset = data.keySet();
				int rownum = 0;
				for (Integer key : keyset) {
					// this creates a new row in the sheet
					if(key==0){
						Row row = sheet.createRow(rownum++);
						int cellnum = 0;
						sheet.setColumnWidth(cellnum,5000);
						Object[] objArr = (Object[]) data.get(key);
						for (Object obj : objArr) {
							// this line creates a cell in the next column of that row
							Cell cell = row.createCell(cellnum++);
							if (obj instanceof String){
								cell.setCellValue((String)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Hstyle);
							}
							else if (obj instanceof Float){
								cell.setCellValue((Float)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Hstyle);
							}
							else if (obj instanceof Integer){
								cell.setCellValue((Integer)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Hstyle);
							}
						}
					}
					else{
						Row row = sheet.createRow(rownum++);
						int cellnum = 0;
						Object[] objArr = (Object[]) data.get(key);
						for (Object obj : objArr) {
							sheet.setColumnWidth(cellnum,5000);
							// this line creates a cell in the next column of that row
							Cell cell = row.createCell(cellnum++);
							if (obj instanceof String){
								cell.setCellValue((String)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Dstyle);
							}
							else if (obj instanceof Float){
								cell.setCellValue((Float)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Dstyle);
							}
							else if (obj instanceof Integer){
								cell.setCellValue((Integer)obj);
								sheet.setColumnWidth(cellnum,5000);
								cell.setCellStyle(Dstyle);
							}
						}
					}
				}
				try {
					//Write on the Excel with the given workspace name
					File file = new File(System.getProperty("user.dir")+"//"+workbookname+".xlsx");
					FileOutputStream out = new FileOutputStream(file);
					workbook.write(out);
					out.close();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
			}
			int no_of_sheet = workbook.getNumberOfSheets();
			System.out.println("no_of_sheet :"+no_of_sheet);
		}
		else {
			workbook = new XSSFWorkbook();

			XSSFCellStyle Dstyle = workbook.createCellStyle();
			//Dstyle.setBorderTop(BorderStyle.DASHED);
			Dstyle.setBorderBottom(BorderStyle.THIN);
			Dstyle.setBorderRight(BorderStyle.THIN);
			Dstyle.setBorderLeft(BorderStyle.THIN);

			XSSFCellStyle Hstyle = workbook.createCellStyle();
			Dstyle.setBorderTop(BorderStyle.THIN);
			Hstyle.setBorderBottom(BorderStyle.THIN);
			Hstyle.setBorderRight(BorderStyle.THIN);
			Hstyle.setBorderLeft(BorderStyle.THIN);
			Hstyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			Hstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

			Font Hfont = workbook.createFont();
			Hfont.setColor(IndexedColors.BLACK.getIndex());
			Hstyle.setFont(Hfont);

			Font Dfont = workbook.createFont();
			Dfont.setColor(IndexedColors.BLACK.getIndex());
			Dstyle.setFont(Dfont);
			// Create a blank sheet
			XSSFSheet sheet = workbook.createSheet(sheetName);

			// Iterate over data and write to sheet
			Set<Integer> keyset = data.keySet();
			int rownum = 0;
			for (Integer key : keyset) {
				// this creates a new row in the sheet
				if(key==0){
					Row row = sheet.createRow(rownum++);
					int cellnum = 0;
					sheet.setColumnWidth(cellnum,5000);
					Object[] objArr = (Object[]) data.get(key);
					for (Object obj : objArr) {
						// this line creates a cell in the next column of that row
						Cell cell = row.createCell(cellnum++);
						if (obj instanceof String){
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellValue((String)obj);
							cell.setCellStyle(Hstyle);
						}
						else if (obj instanceof Float){
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellValue((Float)obj);
							cell.setCellStyle(Hstyle);
						}
						else if (obj instanceof Integer){
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellValue((Integer)obj);
							cell.setCellStyle(Hstyle);
						}
					}
				}
				else{
					Row row = sheet.createRow(rownum++);
					int cellnum = 0;
					sheet.setColumnWidth(cellnum,5000);
					Object[] objArr = (Object[]) data.get(key);
					for (Object obj : objArr) {
						// this line creates a cell in the next column of that row
						Cell cell = row.createCell(cellnum++);
						if (obj instanceof String){
							cell.setCellValue((String)obj);
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellStyle(Dstyle);
						}
						else if (obj instanceof Float){
							cell.setCellValue((Float)obj);
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellStyle(Dstyle);
						}
						else if (obj instanceof Integer){
							cell.setCellValue((Integer)obj);
							sheet.setColumnWidth(cellnum,5000);
							cell.setCellStyle(Dstyle);
						}
					}
				}
			}
			try {
				//Write on the Excel with the given workspace name
				FileOutputStream out = new FileOutputStream(new File(workbookname+".xlsx"));
				workbook.write(out);
				out.close();
			}
			catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	//Getting the result by hitting rest api with access token and parsing the record and putting them in Hashmap
	public static void result(String accessToken) throws IOException, EncryptedDocumentException, InvalidFormatException, ParseException
	{

		String testurl = ul.getData("testurl");
		String StartCreatedDate = System.getProperty("StartCreatedDate");//"2018-05-24T19:24:44.000Z";//"2018-05-07T07:51:43.000Z"
		String EndCreatedDate =System.getProperty("EndCreatedDate");//"2018-05-24T20:00:40.000Z";//"2018-05-07T08:49:41.000Z";
		Date dNow = new Date( );
		SimpleDateFormat fts = new SimpleDateFormat ("dd-MMM HH.mm.ss");
		String ExecutedAt=fts.format(dNow);
		System.out.println("Execution Date & Time: " + ExecutedAt);

		/*SimpleDateFormat ftw = new SimpleDateFormat ("dd-MMM");
		String workbookname=ftw.format(dNow);
		System.out.println("Current Date: " + workbookname);*/
		String str = testurl+"/services/data/v45.0/query?q=SELECT AccountId FROM Case";
		//String str = testurl+"/services/data/v45.0/query?q=SELECT Id,Application,Browser,CreatedDate,EventDate,Platform FROM LoginEvent WHERE EventDate >= 2019-04-01T00:00:00.000Z AND EventDate <= 2019-07-31T23:59:59.000Z";
		//System.out.println("Request String: " + str);
		Response response = null;
		try {
			response = 
					given()
					.auth().oauth2(accessToken)
					.contentType(ContentType.JSON)
					.accept(ContentType.JSON)
					.when()
					.get(str);
		}
		catch(Exception e){
			System.out.println("----------------Not Able To Get Response---------------");
			e.printStackTrace();
		}
		System.out.println(response.prettyPrint());
		System.out.println(str);
		//Json Parser to get attributes
		JsonParser js = new JsonParser();
		JsonObject jsObject = js.parse(response.asString()).getAsJsonObject();
		String totalsize = jsObject.get("totalSize").getAsString();
		System.out.println("Total no. of records :"+totalsize);
		JsonArray jsArray = jsObject.get("records").getAsJsonArray();
		DateFormat dateFormat;
		Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
		data.clear();
		data.put(0, new Object[]{ "ID","Application", "Browser", "Platform","CreatedDate","EventDate"});
		System.out.println("size of array :"+jsArray.size());
		String nextrecordcheck = jsObject.get("done").getAsString();
		JsonObject jarrObject ;
		int i =0;
		int j = i;
		while(nextrecordcheck.equals("false")) {

			for(i=0;i<jsArray.size();i++)
			{
				j=j+1;
				System.out.println("record no:"+j);
				String AccountId = "";
				jarrObject = jsArray.get(i).getAsJsonObject();
				if(jarrObject.get("AccountId").isJsonNull())
				{
					AccountId = null;
				}
				else
				{
					AccountId = jarrObject.get("AccountId").getAsString();
				}

				System.out.println("AccountId : "+AccountId);

				arrlist.add(AccountId);
				//data.put(j, new Object[]{Id,Application,Browser,Platform, date1,date2});
				map.put(String.valueOf(j),new ArrayList<String>(arrlist));
				//}
				//System.out.println();
				arrlist.clear();
			}

			String nextRecordsUrl = jsObject.get("nextRecordsUrl").getAsString();
			String nextrecords = testurl+nextRecordsUrl;
			try {
				response = 
						given()
						.auth().oauth2(accessToken)
						.contentType(ContentType.JSON)
						.accept(ContentType.JSON)
						.when()
						.get(nextrecords);
			}
			catch(Exception e){
				System.out.println("----------------Not Able To Get Second Response---------------");
				e.printStackTrace();
			}

			System.out.println(response.asString());
			JsonParser njs = new JsonParser();
			JsonObject njsObject = njs.parse(response.asString()).getAsJsonObject();
			nextrecordcheck = njsObject.get("done").getAsString();
			jsObject = njsObject;
			jsArray = njsObject.get("records").getAsJsonArray();

		}			
		for(i=0;i<jsArray.size();i++)
		{
			j=j+1;
			System.out.println("record no:"+j);
			jarrObject = jsArray.get(i).getAsJsonObject();
			String AccountId = jarrObject.get("AccountId").getAsString();
			System.out.println("AccountId : "+AccountId);
			arrlist.add(AccountId);
			//data.put(j, new Object[]{Id,Application,Browser,Platform, date1,date2});
			map.put(String.valueOf(j),new ArrayList<String>(arrlist));
			//}

			//System.out.println();

			arrlist.clear();

		}
		String sheetname ="sheet1";
		System.out.println("sheet name is :"+sheetname);
		String workbookname ="BrowserType"; 
		System.out.println("Creating CSV started from EXCEL");
		writeDataAtOnce(map,workbookname);
		//writeExcel(sheetname,workbookname,data);
		//System.out.println("All records are update on Excel Workbook \""+workbookname+"\" and Sheetname \""+sheetname+"\"");
		System.out.println("Creating CSV started from EXCEL");
		//exceltocsv(sheetname,workbookname);
		System.out.println("Creating CSV FINISHED");
	}

	@Test
	public static void check() throws IOException, EncryptedDocumentException, InvalidFormatException, ParseException
	{
		accesstoken =authenticateUser();
		System.out.println(accesstoken);
		result(accesstoken);

	}
}
