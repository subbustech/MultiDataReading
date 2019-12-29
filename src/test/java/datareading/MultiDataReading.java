package datareading;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * 
 * @author Subrahmanyeswara Reddy Padala
 * <br><br>Inspired From QMetry Framework.<br>
 * QMetry Framework has got an excellent way to read data from an excel sheet using @QAFDataProvider
 * annotation.<br>
 * But this can be used only with QMetry framework. Also, I think it lacks setting/updating data in excel.<br>
 * If you want to develop your own framework where you have data driven test cases, you need to 
 * write your own classes for reading and writing data to excel.<br>
 * So, prepared this class using which you can easily read data into a map and also you can set/update data in a cell.
 *
 */

public class MultiDataReading {

	/**
	 * <h1>getData(String path, String sheetName, String key)</h1>
	 * This method returns data from a key range specified in an excel sheet under a workbook.
	 * 
	 * @author Subrahmanyeswara Reddy Padala
	 * @version 1.0
	 * @since 29/12/2019
	 * @param path Excel sheet path Ex: "C:/mydata/data.xlsx" or if the sheet is with in some folder 
	 * like resources in the project under the root folder then "resources/data.xlsx" 
	 * @param sheetName Sheet name where data exists like "Sheet1"
	 * @param key Key for the data. This has to be provided before data (before column names) and after data(after the last cell data.
	 * @return returns a map of integer, map. The second map is a map of String and String. The 
	 * integer is the data row number starting from 1 with in the key range specified. Second map 
	 * contains key value pairs of column name and actual value.
	 */
	public Map<Integer, Map<String, String>> getData(String path, String sheetName, String key) {

		Map<Integer, Map<String, String>> mp = new HashMap<Integer, Map<String, String>>();
		List<String> fieldnames = new ArrayList<String>();

		FileInputStream fis = null;

		try {
			//Get access to the workbook
			fis = new FileInputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		//Create workbook object
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}

		//Get the sheet you want
		XSSFSheet sheet = workbook.getSheet(sheetName);

		//Get all the rows
		Iterator<Row> rows = sheet.iterator();

		boolean initialkeywordcheck = true;

		boolean beginningrowexists = false;
		boolean endingrowexists = false;

		int fieldsrow = 0;
		int databeginningrow = 0;
		int databeginningcell = 0;
		int dataendingrow = 0;
		int dataendingcell = 0;

		while(rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cells = row.iterator();
			while(cells.hasNext()) {
				Cell cell = cells.next();
				if(cell.getCellType()==CellType.STRING)
				{
					if(initialkeywordcheck && cell.getStringCellValue().equalsIgnoreCase(key)) {
						beginningrowexists = true;

						fieldsrow = row.getRowNum();
						databeginningrow = fieldsrow+1;
						databeginningcell = cell.getColumnIndex()+1;

						initialkeywordcheck = false;

						continue;
					}
					if(!initialkeywordcheck && cell.getStringCellValue().equalsIgnoreCase(key)) {
						endingrowexists = true;

						dataendingrow = row.getRowNum();
						dataendingcell = cell.getColumnIndex()-1;

						break;
					}
				}
			}
		}

		if(!beginningrowexists || !endingrowexists) {
			System.out.println("Problem with Keys. Check the keys.");
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if(fieldsrow == dataendingrow) {
			System.out.println("Keys should not be on the same row");
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		//Fetching the column headings into a list
		Row columnheaders = sheet.getRow(fieldsrow);
		Iterator<Cell> columnnames = columnheaders.iterator();
		columnnames.next();
		while(columnnames.hasNext()) {
			Cell cell = columnnames.next();
			fieldnames.add(cell.getStringCellValue());
		}

		//Adding column headings and data to map
		int datarowcount = 1;
		for(int i = databeginningrow; i<=dataendingrow; i++) {
			Row row = sheet.getRow(i);
			Map<String, String> indrowdata = new HashMap<String, String>();
			int fieldcellcount = 0;
			for(int j=databeginningcell; j<=dataendingcell; j++) {
				Cell cell = row.getCell(j);
				DataFormatter df = new DataFormatter();
				String cellvalue = df.formatCellValue(cell);
				indrowdata.put(fieldnames.get(fieldcellcount), cellvalue);
				fieldcellcount +=1;
			}
			mp.put(datarowcount, indrowdata);
			datarowcount += 1;
		}

		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return mp;
	}

	/**
	 * <h1>getData(String path, String sheetName)</h1>
	 * This method returns data from an excel sheet under a workbook.
	 * 
	 * @author Subrahmanyeswara Reddy Padala
	 * @version 1.0
	 * @since 29/12/2019
	 * @param path Excel sheet path Ex: "C:/mydata/data.xlsx" or if the sheet is with in some folder 
	 * like resources in the project under the root folder then "resources/data.xlsx" 
	 * @param sheetName Sheet name where data exists like "Sheet1"
	 * @return returns a map of integer, map. The second map is a map of String and String. The 
	 * integer is the data row number starting from 1. Second map contains key value pairs of 
	 * column name and actual value.
	 */
	public Map<Integer, Map<String, String>> getData(String path, String sheetName) {

		Map<Integer, Map<String, String>> mp = new HashMap<Integer, Map<String, String>>();
		List<String> fieldNames = new ArrayList<String>();

		FileInputStream fis = null;
		try {
			//Get access to the workbook
			fis = new FileInputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		//Create workbook object
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}

		//Get the sheet you want
		XSSFSheet sheet = workbook.getSheet(sheetName);

		//Get all the column names into a list
		Row fieldrows = sheet.getRow(0);
		Iterator<Cell> fieldcolumns = fieldrows.iterator();
		while(fieldcolumns.hasNext()) {
			Cell cell = (Cell) fieldcolumns.next();
			DataFormatter df = new DataFormatter();
			String cellvalue = df.formatCellValue(cell);
			fieldNames.add(cellvalue);
		}

		int databeginningrow = 1;
		int databeginningcell = 0;
		int dataendingrow = sheet.getLastRowNum();
		System.out.println("dataendingrow"+dataendingrow);
		int dataendingcell = fieldrows.getLastCellNum()-1;
		System.out.println("lastcellnum"+dataendingcell);

		//Adding column headings and data to map
		int datarowcount = 1;
		for(int i = databeginningrow; i<=dataendingrow; i++) {
			Row row = sheet.getRow(i);
			Map<String, String> indrowdata = new HashMap<String, String>();
			int fieldcellcount = 0;
			for(int j=databeginningcell; j<=dataendingcell; j++) {
				Cell cell = row.getCell(j);
				DataFormatter df = new DataFormatter();
				String cellvalue = df.formatCellValue(cell);
				indrowdata.put(fieldNames.get(fieldcellcount), cellvalue);
				fieldcellcount +=1;
			}
			mp.put(datarowcount, indrowdata);
			datarowcount += 1;
		}

		System.out.println(mp);
		System.out.println(mp.size());
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return mp;
	}


	/**
	 * <h1>setData(String path, String sheetName, String key, int rowno, String colName, String value)</h1>
	 * This method updates a particular cell data with in a key range specified in an excel sheet under a workbook.
	 * 
	 * @author Subrahmanyeswara Reddy Padala
	 * @version 1.0
	 * @since 29/12/2019
	 * @param path Excel sheet path Ex: "C:/mydata/data.xlsx" or if the sheet is with in some folder 
	 * like resources in the project under the root folder then "resources/data.xlsx" 
	 * @param sheetName Sheet name where data exists like "Sheet1"
	 * @param key Key for the data. This has to be provided before data (before column names) and after data(after the last cell data.
	 * @param rowno This is the data row number (not including the column names) starting from 1 with in the key range.
	 * @param colName Name of the column where the cell has to be updated.
	 * @param value Value that has to be updated.
	 * @return returns boolean value true, if the data is updated successfully.
	 */
	public boolean setData(String path, String sheetName, String key, int rowno, String colName, String value) {

		FileInputStream fis = null;

		try {
			//Get access to the workbook
			fis = new FileInputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		//Create workbook object
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}

		//Get the sheet you want
		XSSFSheet sheet = workbook.getSheet(sheetName);

		//Get all the rows
		Iterator<Row> rows = sheet.iterator();

		int fieldsrow = 0;
		int actualrow = 0;
		int actualcol = 0;
		boolean rowfound = false;
		boolean colfound = false;
		while(rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cells = row.iterator();
			while(cells.hasNext()) {
				Cell cell = cells.next();
				if(cell.getCellType()==CellType.STRING)
				{
					if(cell.getStringCellValue().equalsIgnoreCase(key)) {

						fieldsrow = row.getRowNum();
						actualrow = fieldsrow + rowno;
						rowfound = true;
					}
				}
				if(cell.getCellType()==CellType.STRING)
				{
					if(cell.getStringCellValue().equalsIgnoreCase(colName)) {

						actualcol = cell.getColumnIndex();
						colfound = true;
					}
				}
				if(rowfound && colfound) {
					break;
				}
			}
		}
		//		System.out.println("actrow "+actualrow);
		//		System.out.println("actcol "+actualcol);

		try {
			fis.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		FileOutputStream fos = null;

		try {
			//Get access to the workbook
			fos = new FileOutputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		Row rowtoput = sheet.getRow(actualrow);
		Cell celltoput = rowtoput.createCell(actualcol);
		celltoput.setCellValue(value);
		boolean success = false;
		try {
			workbook.write(fos);
			success=true;
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		try {
			fos.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return success;
	}

	/**
	 * <h1>setData(String path, String sheetName, int rowno, String colName, String value)</h1>
	 * This method updates a particular cell data in an excel sheet under a workbook.
	 * 
	 * @author Subrahmanyeswara Reddy Padala
	 * @version 1.0
	 * @since 29/12/2019
	 * @param path Excel sheet path Ex: "C:/mydata/data.xlsx" or if the sheet is with in some folder 
	 * like resources in the project under the root folder then "resources/data.xlsx" 
	 * @param sheetName Sheet name where data exists like "Sheet1"
	 * @param rowno This is the data row number (not including the column names) starting from 1 with in the key range.
	 * @param colName Name of the column where the cell has to be updated.
	 * @param value Value that has to be updated.
	 * @return returns boolean value true, if the data is updated successfully.
	 */
	public boolean setData(String path, String sheetName, int rowno, String colName, String value) {

		FileInputStream fis = null;

		try {
			//Get access to the workbook
			fis = new FileInputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		//Create workbook object
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}

		//Get the sheet you want
		XSSFSheet sheet = workbook.getSheet(sheetName);

		//Get the column names row
		Row row = sheet.getRow(0);

		int actualcol = 0;
		boolean colfound = false;

		Iterator<Cell> cells = row.iterator();
		while(cells.hasNext()) {
			Cell cell = cells.next();
			if(cell.getCellType()==CellType.STRING)
			{
				if(cell.getStringCellValue().equalsIgnoreCase(colName)) {
					actualcol = cell.getColumnIndex();
					colfound = true;
				}
			}
			if(colfound) {
				break;
			}
		}

		try {
			fis.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		FileOutputStream fos = null;

		try {
			//Get access to the workbook
			fos = new FileOutputStream(path);
		}
		catch(IOException e){
			System.out.println(e.getMessage());
		}

		Row rowtoput = sheet.getRow(rowno);
		Cell celltoput = rowtoput.createCell(actualcol);
		celltoput.setCellValue(value);
		boolean success = false;
		try {
			workbook.write(fos);
			success=true;
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		try {
			fos.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}

		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return success;
	}
}
