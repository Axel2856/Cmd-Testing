package Demo;

import org.testng.annotations.Test;

import ExcelUtility.ExcelColorUtil;

public class ExcelColorConcept {

	String outputpath="E:\\Selenium_Evengbatch\\CmdTestng\\ExcelOutput\\ColorResult.xlsx";
	@Test
	public void setData()throws Throwable
	{
		ExcelColorUtil xl=new ExcelColorUtil("E:\\Selenium_Evengbatch\\CmdTestng\\ExcelInput\\ExcelColor.xlsx");
		xl.setCellData("Data1", 1, "Result", "Axel", outputpath);
		//xl.setCellData("Data1", 2, "Result", "Blaze", outputpath);
	}
}
