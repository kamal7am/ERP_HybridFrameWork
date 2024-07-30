package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLfile_Util {
	XSSFWorkbook wb;
	public XLfile_Util(String Excelsheet ) throws Throwable {
		FileInputStream fi= new FileInputStream(Excelsheet);
		wb= new XSSFWorkbook(fi);
	}

	//method for row count
	public int rowcount(String sheetname ) {

		return wb.getSheet(sheetname).getLastRowNum();
	}
	public String getcelldata(String sheetname,int row,int column) {
		String data="";
		if(wb.getSheet(sheetname).getRow(row).getCell(column).getCellType()==CellType.NUMERIC) {
			int celldata=(int) wb.getSheet(sheetname).getRow(row).getCell(column).getNumericCellValue();
			data= String.valueOf(celldata);
		}
		else {
			data=wb.getSheet(sheetname).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}

	//method for get cell data
	public void setcelldata(String sheetname,int row,int column,String Status,String outputpath) throws Throwable {
		XSSFSheet sheet=wb.getSheet(sheetname);
		XSSFRow rownum=sheet.getRow(row);
		XSSFCell cell=rownum.createCell(column);
		cell.setCellValue(Status);
		if (Status.equalsIgnoreCase("PASS")) {
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.GREEN.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if (Status.equalsIgnoreCase("FAIL")) {
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			font.setColor(IndexedColors.RED.getIndex());
			font.setBold(true);
			style.setFont(font);
			rownum.getCell(column).setCellStyle(style);
		}
		else if (Status .equalsIgnoreCase("Blocked")) {
			XSSFCellStyle style=wb.createCellStyle();
			XSSFFont font=wb.createFont();
			style.setFont(font);
			font.setColor(IndexedColors.BLUE.getIndex());
			font.setBold(true);
			rownum.getCell(column).setCellStyle(style);
		}
		FileOutputStream fo=new FileOutputStream(outputpath);
		wb.write(fo);
	}

}
