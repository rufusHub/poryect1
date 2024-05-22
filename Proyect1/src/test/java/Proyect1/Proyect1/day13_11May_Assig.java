package Proyect1.Proyect1;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class day13_11May_Assig {
	public void ReadDataBasedUponRowNoAndColumnNo(String Path, int row, int col) throws BiffException, IOException {
		File f = new File(Path);
		Workbook wk = Workbook.getWorkbook(f);
		Sheet ws = wk.getSheet(0);
		int r = ws.getRows();
		int c = ws.getColumns();
		for(int i=0 ; i<r ; ++i) {			// loop for rows.
			
			if (i != row) {continue;}
			else {
				for(int j=0 ; j<c ; ++j) {		// loop for columns.
					if (j != col) {continue;}
					else {
						Cell c1 = ws.getCell(j, i);	// (col, row)
						System.out.println(c1.getContents());
					}
				}	
			}
		}
	}
	
	public void ReadDataBaedUponRowNo(String Path, int row) throws BiffException, IOException {
		File f = new File(Path);
		Workbook wk = Workbook.getWorkbook(f);
		Sheet ws = wk.getSheet(0);
		int r = ws.getRows();
		int c = ws.getColumns();
		for(int i=0 ; i<r ; ++i) {			    // loop for rows.
			if (i == row) {
				for(int j=0 ; j<c ; ++j) {		// loop for columns.	
					Cell c1 = ws.getCell(j, i);	// (col, row)
					System.out.println(c1.getContents());
				}
			}
			else {continue;}		
		}
	}
	
	public void ReadDataBasedUponRange(String Path, int row_ini, int row_end) throws BiffException, IOException {
		File f = new File(Path);
		Workbook wk = Workbook.getWorkbook(f);
		Sheet ws = wk.getSheet(0);
		int r = ws.getRows();
		int c = ws.getColumns();
		for(int i=row_ini-1 ; i<row_end ; ++i) {			// loop for rows.
			for(int j=0 ; j<c ; ++j) {		// loop for columns.
				Cell c1 = ws.getCell(j, i);	// (col, row)
				System.out.println(c1.getContents());
			}	
		}
	}
	
	public void WriteData(String Path, int row, int col) throws IOException, RowsExceededException, WriteException {
		File f = new File(Path);	//connection
		WritableWorkbook wk = Workbook.createWorkbook(f);
		WritableSheet ws = wk.createSheet("Rodrigo", 0);
		for(int i=0 ; i<row ; ++i) {		//loop for rows
			for(int j=0 ; j<col ; ++j) {  //loop for columns
				System.out.println("(row:" + i + " col:" + j + ") .Enter to values x : ");
				Scanner ob = new Scanner(System.in);
				String x = ob.next();
				Label L = new Label(j,i, x);
				ws.addCell(L);
			}
		}
		wk.write();
		wk.close();
	}
	
	public void CopyPaste(String Path1, String Path2) throws BiffException, IOException, RowsExceededException, WriteException {
		
		File f = new File(Path1);				// Read
		Workbook wk = Workbook.getWorkbook(f);
		Sheet ws = wk.getSheet(0);
		int r = ws.getRows();
		int c = ws.getColumns();
		
		File f2 = new File(Path2);	//connection
		WritableWorkbook wk2 = Workbook.createWorkbook(f2);
		WritableSheet ws2 = wk2.createSheet("Rodrigo", 0);

		for(int i=0 ; i<r ; ++i) {			// loop for rows.			
			for(int j=0 ; j<c ; ++j) {		// loop for columns.
				Cell c1 = ws.getCell(j, i);	// (col, row)
				Label L = new Label(j,i,(String)c1.getContents());
				ws2.addCell(L);
			}	
		}	
		wk2.write();
		wk2.close();
		
	}
	
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException {
		day13_11May_Assig obj = new day13_11May_Assig();
		String Path = "../Proyect1/Book1.xls";
		String Path3 = "../Proyect1/Book3.xls";
//		obj.ReadDataBasedUponRowNoAndColumnNo(Path, 3, 3);
//		obj.ReadDataBaedUponRowNo(Path, 4);
//		obj.ReadDataBasedUponRange(Path, 4, 6);
//		obj.WriteData(Path3, 2, 2 );
		obj.CopyPaste(Path, Path3);
	}
}
