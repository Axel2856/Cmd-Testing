package ExcelUtility;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelColorUtil {
	FileInputStream fi;
	FileOutputStream fo;
	Workbook wb;
	Sheet ws;
	Row rownum;
	Cell cell;

	public ExcelColorUtil(String excelpath) throws Throwable {
		fi = new FileInputStream(excelpath);
		wb = WorkbookFactory.create(fi);
	}

	public int rowCount(String sheetname) {
		return wb.getSheet(sheetname).getLastRowNum();
	}

	public int colCount(String sheetname) {
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}

	public String getCellData(String sheetname, int row, String columnname) {
		String data = "";
		for (int i = 0; i < wb.getSheet(sheetname).getRow(0).getLastCellNum(); i++) {
			String match = wb.getSheet(sheetname).getRow(0).getCell(i).getStringCellValue();
			if (match.equalsIgnoreCase(columnname)) {
				data = wb.getSheet(sheetname).getRow(row).getCell(i).getStringCellValue();
				// System.out.println("The column Name is:="+match);
			}
		}
		return data;
	}

	// Set Data in Cell of Excel Sheet Method
	public void setCellData(String sheetname, int row, String columnname, String status, String outputexcel)
			throws Throwable {
		for (int i = 0; i < wb.getSheet(sheetname).getRow(0).getLastCellNum(); i++) {
			String match = wb.getSheet(sheetname).getRow(0).getCell(i).getStringCellValue();
			if (match.equalsIgnoreCase(columnname)) {
				ws = wb.getSheet(sheetname);
				rownum = ws.getRow(row);
				cell = rownum.createCell(i);
				cell.setCellValue(status);
				
				
				CellStyle style=wb.createCellStyle();
				//use either foreground or background , don't use both(we can use but not cleared any color,only foreground visible)
				//GREY_25_PERCENT//MAROON//DARK_BLUE
				style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());				
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				
				//style.setFillBackgroundColor(IndexedColors.DARK_BLUE.getIndex());
				//style.setFillPattern(FillPatternType.BIG_SPOTS);
				
				Font font=wb.createFont();
				font.setBoldweight(Font.BOLDWEIGHT_BOLD);
				font.setBold(true);
				font.setColor(IndexedColors.MAROON.getIndex());
				
				/////////////font.setFontHeight((short)12);
				
											
				
				//Below codes Working Perfectly
				//font.setFontHeightInPoints((short)12);  
	            //font.setFontName("Arial Black");  
	            //font.setItalic(true);  
	            //font.setStrikeout(true);  
				
				style.setFont(font);	
				cell.setCellStyle(style);
				//rownum.getCell(i).setCellStyle(style);//i stands for column index number 
				
				fo = new FileOutputStream(outputexcel);
				wb.write(fo);
			}

		}
	}

}
