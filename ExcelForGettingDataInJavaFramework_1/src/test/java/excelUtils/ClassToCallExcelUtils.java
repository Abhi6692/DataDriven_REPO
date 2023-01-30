package excelUtils;

public class ClassToCallExcelUtils {

	public static void main(String[] args) {
		
		
		String projectPath = System.getProperty("user.dir");
		ExcelUtilities excelUtils = new ExcelUtilities(projectPath +"\\src\\test\\resources\\TestData.xlsx" , "sheet1");
		
		excelUtils.getRowCount();
		excelUtils.getColCount();
		excelUtils.getStringCellData(0, 0);
		excelUtils.getNumericCellData(1, 1);
		
	
	}

}
