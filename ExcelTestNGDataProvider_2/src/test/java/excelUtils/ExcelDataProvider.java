package excelUtils;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProvider {
	

	/**
	 * This is the test method which pulls the data provider method
	 * @param UserName
	 * @param Password
	 * @param Company
	 */
	@Test(dataProvider ="GetData" )
	public void testMethod(String UserName , String Password , String Company) {
		
		System.out.println( "The Username is --> " + UserName + " | " + "The Password is --> " + Password+ " | " + "The company is -->  " + Company );
	}
	
	
	/**
	 * This is the Data Provider method
	 * @return
	 */
	@DataProvider(name = "GetData")
	public Object[][] getData() {
		
		String excelPath = "C:\\Users\\OM\\Desktop\\Edureka\\ExcelTestNGDataProvider_2\\src\\test\\resources\\TestData.xlsx";
		return testData(excelPath, "sheet1");
		
		
	}
	
	/**
	 * This method is created to call this methods inside the dataprovider method
	 * @param excelPath
	 * @param sheetName
	 * @return
	 */
	public static Object[][] testData(String excelPath , String sheetName ) {
		
		ExcelUtilities excelutility = new ExcelUtilities(excelPath, sheetName);
		
		int rowCount = excelutility.getRowCount();
		int colCount = excelutility.getColCount();
		
		Object data[][] = new Object[rowCount-1][colCount];
		
		 
		
		//Starting from index 1 because index '0' is the header row
		for(int i = 1 ; i<rowCount ; i++) {
			
			for(int j = 0; j<colCount; j++) {
			
				String cellData = excelutility.getStringCellData(i, j);	
				data[i-1][j] = cellData;
				
			}
		}
		
		return data;
	}
	
}
