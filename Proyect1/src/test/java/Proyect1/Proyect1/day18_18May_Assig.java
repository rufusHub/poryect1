package Proyect1.Proyect1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class day18_18May_Assig {

	public void ReadDataBasedUponRowNoAndColumnNo(String path, int col, int row) throws IOException {
		File f = new File(path);						// file object
		FileInputStream fi = new FileInputStream(f);	// inputstream object
		XSSFWorkbook xs = new XSSFWorkbook(fi);			// workbook object
		XSSFSheet xt = xs.getSheetAt(0);				// sheet object
		int r = xt.getPhysicalNumberOfRows();			// fetch the # of rows
		for(int i=0; i<r; ++i) {						// loop for rows
			if(i != row) {continue;}
			else {
				XSSFRow xr = xt.getRow(i);					// row object
				int c = xr.getPhysicalNumberOfCells();		// fetch # of columns
				for(int j=0; j<c; ++j) {					// loop for columns
						if(j != col) {continue;}
						else {
							XSSFCell xc = xr.getCell(j);			// cell object
							System.out.println(xc.getStringCellValue());	// get cell value
						}
				}
			}
		}	
	}
	
	public void ReadDataBasedUponRowNo(String path, int row) throws IOException {
		File f = new File(path);						// file object
		FileInputStream fi = new FileInputStream(f);	// inputstream object
		XSSFWorkbook xs = new XSSFWorkbook(fi);			// workbook object
		XSSFSheet xt = xs.getSheetAt(0);				// sheet object
		int r = xt.getPhysicalNumberOfRows();			// fetch the # of rows
		for(int i=0; i<r; ++i) {						// loop for rows
			if(i == row) {
				XSSFRow xr = xt.getRow(i);					// row object
				int c = xr.getPhysicalNumberOfCells();		// fetch # of columns
				for(int j=0; j<c; ++j) {					// loop for columns
					XSSFCell xc = xr.getCell(j);			// cell object
					System.out.println(xc.getStringCellValue());	// get cell value
				}
			}
			else {continue;}
		}
	}

	public void ReadDataBasedUponRange(String path, int row_ini, int row_end ) throws IOException {
		File f = new File(path);						// file object
		FileInputStream fi = new FileInputStream(f);	// inputstream object
		XSSFWorkbook xs = new XSSFWorkbook(fi);			// workbook object
		XSSFSheet xt = xs.getSheetAt(0);				// sheet object
		int r = xt.getPhysicalNumberOfRows();			// fetch the # of rows
		for(int i=row_ini; i<row_end; ++i) {						// loop for rows
			XSSFRow xr = xt.getRow(i);					// row object
			int c = xr.getPhysicalNumberOfCells();		// fetch # of columns
			for(int j=0; j<c; ++j) {					// loop for columns
				XSSFCell xc = xr.getCell(j);			// cell object
				System.out.println(xc.getStringCellValue());	// get cell value
			}
		}
	}

	public void WriteData(String path, int row, int col) throws IOException {
		File f = new File(path);
		FileOutputStream fo =  new FileOutputStream(f);	// file object.
		XSSFWorkbook xs = new XSSFWorkbook();			// workbook object
		XSSFSheet xt = xs.createSheet("SHEET1");		// sheet object
		for (int i=0; i<row; ++i) {						// loop for rows
			XSSFRow xr = xt.createRow(i);				// row object.
			for(int j=0; j<col; ++j) {					// loop for columns.
				XSSFCell xc = xr.createCell(j);			// column object.
				System.out.println("(row:" + i + " col:" + j + ") .Enter to values x : ");
				Scanner ob = new Scanner(System.in);
				String x = ob.next();
				xc.setCellValue(x);				// set the cell data.
			}
		}
		xs.write(fo); // move data from workbook to the stream.
		fo.flush();	  // move the data from stream to file.
		fo.close();
	}

	public void CopyPaste(String path1, String path2) throws IOException {
		File f1 = new File(path1);						// file object
		FileInputStream fi = new FileInputStream(f1);	// inputstream object
		XSSFWorkbook xs1 = new XSSFWorkbook(fi);			// workbook object
		XSSFSheet xt1 = xs1.getSheetAt(0);				// sheet object
		
		File f2 = new File(path2);
		FileOutputStream fo =  new FileOutputStream(f2);	// file object.
		XSSFWorkbook xs2 = new XSSFWorkbook();			// workbook object
		XSSFSheet xt2 = xs2.createSheet("SHEET1");		// sheet object
		
		int r = xt1.getPhysicalNumberOfRows();			// fetch the # of rows
		for(int i=0; i<r; ++i) {						// loop for rows
			XSSFRow xr1 = xt1.getRow(i);					// row object
			int c = xr1.getPhysicalNumberOfCells();		// fetch # of columns
			XSSFRow xr2 = xt2.createRow(i);				// row object.
			for(int j=0; j<c; ++j) {					// loop for columns
				XSSFCell xc1 = xr1.getCell(j);			// cell object
				XSSFCell xc2 = xr2.createCell(j);			// column object.
				//System.out.println(xc1.getStringCellValue());	// get cell value
				xc2.setCellValue(xc1.getStringCellValue());				// set the cell data.
			}
		}
		
		xs2.write(fo); // move data from workbook to the stream.
		fo.flush();	  // move the data from stream to file.
		fo.close();	
	}
	
	
	public static void main(String[] args) throws IOException {
		String path1 = "../Proyect1/Book5.xlsx";
		String path2 = "../Proyect1/Book4.xlsx";
		day18_18May_Assig obj = new day18_18May_Assig();
//		obj.ReadDataBasedUponRowNoAndColumnNo(path1, 3, 3);
//		obj.ReadDataBasedUponRowNo(path1, 3);
//		obj.ReadDataBasedUponRange(path1, 0, 3);
//		obj.WriteData(path2, 3, 3);
		obj.CopyPaste(path1, path2);
	}
	
	
	
}
