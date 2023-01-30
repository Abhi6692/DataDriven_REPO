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
	 * @return 
	 */
	public  int getRowCount() {

		int rowCount = 0;
		try {

			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("Number of rows = " +rowCount );

		} 


		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}

		return rowCount;
	}



	/**
	 * This method is for fetching the actual column count of the excel having data
	 * @return 
	 */
	public int getColCount() {

		int colCount =0;
		try {

			colCount = sheet.getRow(0).getPhysicalNumberOfCells();
			System.out.println("Number of columns = " +colCount );

		} 


		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return  colCount;
	}





	/**
	 * This method is is used to get the String cell data from excel
	 * @return 
	 */
	public  String getStringCellData(int rowNum , int colNum) {

		String cellData = null;
		try {
			cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println("The String cell data is " +cellData  );

		} 

		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return cellData;
	}


	/**
	 * This method is is used to get the Numeric cell data from excel
	 * @return 
	 */
	public  double getNumericCellData(int rowNum , int colNum) {
		double celldata_Numeric = 0;

		try {
			celldata_Numeric = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println("The Numeric cell data is " +celldata_Numeric);

		} 

		catch (Exception e) {
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();
		}
		return celldata_Numeric;
	}



}
