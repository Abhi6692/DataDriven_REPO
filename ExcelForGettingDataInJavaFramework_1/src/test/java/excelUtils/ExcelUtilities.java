package excelUtils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtilities {

	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	//Constructor
	public ExcelUtilities(String excelPath, String sheetNamne){
	
		try {
			
			
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetNamne);// can be used with sheet name or index
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		
	}
	
	/**
	 * This method is for fetching the actual row count of the excel having data
	 */
	public static void getRowCount() {

		try {

			int rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("Number of rows = " +rowCount );

		} 


		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}

	}

	
	
	/**
	 * This method is for fetching the actual column count of the excel having data
	 */
	public static void getColCount() {

		try {

			int colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("Number of columns = " +colCount );

		} 


		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}

	}
	
	
	
	

	/**
	 * This method is is used to get the String cell data from excel
	 */
	public static void getStringCellData(int rowNum , int colNum) {

		try {
			String cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println("The String cell data is " +cellData  );
			
		} 

		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}

	}


	/**
	 * This method is is used to get the Numeric cell data from excel
	 */
	public static void getNumericCellData(int rowNum , int colNum) {

		try {
			double celldata_Numeric = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println("The Numeric cell data is " +celldata_Numeric);

		} 

		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}

	}

	
	
}
