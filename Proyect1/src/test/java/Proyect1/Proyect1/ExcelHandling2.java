package Proyect1.Proyect1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelHandling2 {
	// Read excel file xlsx with Apache poi library.
	public void readExcelData(String path) throws IOException {
		File f = new File(path);						// file object
		FileInputStream fi = new FileInputStream(f);	// inputstream object
		XSSFWorkbook xs = new XSSFWorkbook(fi);			// workbook object
		XSSFSheet xt = xs.getSheetAt(0);				// sheet object
		int r = xt.getPhysicalNumberOfRows();			// fetch the # of rows
		for(int i=0; i<r; ++i) {						// loop for rows
			XSSFRow xr = xt.getRow(i);					// row object
			int c = xr.getPhysicalNumberOfCells();		// fetch # of columns
			for(int j=0; j<c; ++j) {					// loop for columns
				XSSFCell xc = xr.getCell(j);			// cell object
				System.out.println(xc.getStringCellValue());	// get cell value
			}
		}
	}
	
	
	// Write excel file xlsx with Apache poi library.
	public void writeExcelData(String path) throws IOException {
		File f = new File(path);
		FileOutputStream fo =  new FileOutputStream(f);	// file object.
		XSSFWorkbook xs = new XSSFWorkbook();			// workbook object
		XSSFSheet xt = xs.createSheet("SHEET1");		// sheet object
		for (int i=0; i<3; ++i) {						// loop for rows
			XSSFRow xr = xt.createRow(i);				// row object.
			for(int j=0; j<3; ++j) {					// loop for columns.
				XSSFCell xc = xr.createCell(j);			// column object.
				xc.setCellValue("Rodrigo");				// set the cell data.
			}
		}
		xs.write(fo); // move data from workbook to the stream.
		fo.flush();	  // move the data from stream to file.
		fo.close();
	}
	
	public static void main(String[] args) throws IOException {
		String path = "../Proyect1/Book4.xlsx"; 
		ExcelHandling2 e = new ExcelHandling2();
//		e.writeExcelData(path);
		e.readExcelData(path);
	}
	
	
	
}
