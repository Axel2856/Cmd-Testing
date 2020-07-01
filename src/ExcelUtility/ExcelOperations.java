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

public class ExcelOperations {
	FileInputStream fi;
	FileOutputStream fo;
	Workbook wb;
	Sheet ws;
	Row rownum;
	Cell cell;

	// Constructor Set Method For Access Excel Sheet
	public ExcelOperations(String excelpath) throws Throwable {
		fi = new FileInputStream(excelpath);
		wb = WorkbookFactory.create(fi);
	}

	// find whether sheets exists
	public void isSheetExist(String sheetname) {
		int index = wb.getSheetIndex(sheetname);
		if (index == -1) {
			System.out.println("The Entered Sheetname (" + sheetname + ") is not Exist");
		} else {
			System.out.println("The Entered Sheetname (" + sheetname + ") is Exist");
		}
	}

	// Row Count Method
	public int rowCount(String sheetname) {
		return wb.getSheet(sheetname).getLastRowNum();
	}

	// Column Count Method
	public int colCount(String sheetname) {
		return wb.getSheet(sheetname).getRow(0).getLastCellNum();
	}

	// Cell Data Retrieve from Excel Sheet Method
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
				fo = new FileOutputStream(outputexcel);
				wb.write(fo);
			}

		}
	}

	// For New Sheet Creation in Workbook
	public boolean addSheet(String sheetname, String outputexcel) {
		try {
			wb.createSheet(sheetname);
			fo = new FileOutputStream(outputexcel);
			wb.write(fo);
			fo.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// Remove Sheet From Workbook
	public boolean removeSheet(String sheetname, String outputexcel) {
		int index = wb.getSheetIndex(sheetname);
		if (index == -1)
			return false;
		try {
			wb.removeSheetAt(index);
			fo = new FileOutputStream(outputexcel);
			wb.write(fo);
			fo.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// Add New Column in Excel Sheet
	public boolean addColumn(String sheetname, String columnname, String outputexcel) throws Throwable {
		try {
			int index = wb.getSheetIndex(sheetname);
			if (index == -1) {
				System.out.println("No Sheet Found,Entered Sheet Name is:=" + sheetname);
				return false;
			}
			ws = wb.getSheetAt(index);

			rownum = ws.getRow(0);
			if (rownum == null)
				rownum = ws.createRow(0);

			// For Color the Cell and Column Name Below Code
			CellStyle style = wb.createCellStyle();
			style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			Font font = wb.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			font.setBold(true);
			font.setColor(IndexedColors.MAROON.getIndex());
			style.setFont(font);
			// cell.setCellStyle(style);
			
			//below code for check column exist or not
			if (rownum.getLastCellNum() >= 0) {
				int totalcolumn = wb.getSheet(sheetname).getRow(0).getLastCellNum();
				for (int i = 0; i < totalcolumn; i++) {
					String match = wb.getSheet(sheetname).getRow(0).getCell(i).getStringCellValue();
					if (match.equalsIgnoreCase(columnname)) {
						System.out.println(
								"Column(" + match + ")" + " already present inside the SheetName:=" + sheetname);
						break;
					} else if (!(match.equalsIgnoreCase(columnname)) && i == totalcolumn - 1) {
						cell = rownum.createCell(rownum.getLastCellNum());
						cell.setCellStyle(style);
						cell.setCellValue(columnname);
					}

				} // for closed
			} // if(rownum.getLastCellNum() >= 0) closed
			if (rownum.getLastCellNum() == -1) {
				cell = rownum.createCell(0);
				cell.setCellStyle(style);
				cell.setCellValue(columnname);
			}

			fo = new FileOutputStream(outputexcel);
			wb.write(fo);
			fo.close();
			// } // index else closed
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}// addColumn() closed
	
	// Remove Column From Existing Sheet in Workbook
	public boolean removeColumn(String sheetname, String columnname, String outputexcel) throws Throwable {
		try {
			int index = wb.getSheetIndex(sheetname);
			if (index == -1) {
				System.out.println("No Sheet Found,Entered Sheet Name is:=" + sheetname);
				return false;
			}

			else {
				int totalcolumn = wb.getSheet(sheetname).getRow(0).getLastCellNum();
				for (int i = 0; i < totalcolumn; i++) {
					String match = wb.getSheet(sheetname).getRow(0).getCell(i).getStringCellValue();
					if (match.equalsIgnoreCase(columnname)) {
						// j loop for fetch each row data
						for (int j = 0; j <= wb.getSheet(sheetname).getLastRowNum(); j++) {
							rownum = ws.getRow(j);
							if (rownum != null) {
								cell = rownum.getCell(i);
								if (cell != null) {
									rownum.removeCell(cell);
								}
							} // rownum if
						} // j for close
					}
				} // i for closed
				fo = new FileOutputStream(outputexcel);
				wb.write(fo);
				fo.close();
			} // else closed
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}// Remove Column Method closed
}


//For Coloring Cell and font also

/*//Set Data in Cell of Excel Sheet Method
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
*/	