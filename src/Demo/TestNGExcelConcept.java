package Demo;

import org.testng.Assert;
import org.testng.annotations.Test;

import ExcelUtility.ExcelOperations;

public class TestNGExcelConcept {
	String inputpath="E:\\Selenium_Evengbatch\\CmdTestng\\ExcelInput\\ExcelEmpDemo.xlsx";
	String outputpath="E:\\Selenium_Evengbatch\\CmdTestng\\ExcelOutput\\AddColExcel.xlsx";
	@Test
	public void testMethod()throws Throwable {
		ExcelOperations read=new ExcelOperations(inputpath);
		//We can enter Sheetname any type means it's not case sensitive like(Login=login=LogIN=loGiN)
		int rc=read.rowCount("login");		
		int cc=read.colCount("Login");
		System.out.println("Total Number of Rows is:="+rc+" Total Number of Columns is:="+cc);
		//getCellData() calling to fetch value from Excel file
		String celldata=read.getCellData("Login", 2, "password");
		System.out.println("The cell value is:="+celldata);
		/*for(int j=1;j<=rc;j++)
		{
			String celldata=read.getCellData("Login", j, "password");
			System.out.println("The cell value is:="+celldata);
		}*/
		//setCellData() calling to set data in Excel File
		//read.setCellData("Login", 1, "Status", "Pass", outputpath);
		//Create New Sheet in ExcelSheet
		//read.addSheet("NewSheet", outputpath);
		//Remove Sheet From Excel Sheet
		//read.removeSheet("NewSheet", outputpath);
		//Add Column to Existing Sheet in Workbook
		//read.addColumn("Login", "test", outputpath);
		//Remove Column and entire data from Existing Sheet in Workbook
		//read.removeColumn("Login", "password", outputpath);
		// find whether sheets exists
		//read.isSheetExist("login");
		
		//boolean res=read.addSheet("NewSheet", outputpath);
		//Assert.assertTrue(res);
		//Thread.sleep(5000);
		//read.isSheetExist("NewSheet");
		//read.addColumn("NewSheet","KEYS",outputpath);
		//read.addColumn("NewSheet", "VALUES", outputpath);
		//Thread.sleep(3000);
		read.addColumn("Login","PASSWORD",outputpath);		
		read.addColumn("Login", "STATUS", outputpath);
		read.addColumn("Login", "RESULTS", outputpath);
	}
}
