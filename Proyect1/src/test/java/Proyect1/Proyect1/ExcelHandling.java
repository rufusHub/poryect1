package Proyect1.Proyect1;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelHandling {
	
	// Write excel file 'xls' with 'JXL' library.
	public void writeData(String Path) throws IOException, RowsExceededException, WriteException {
		File f = new File(Path);	//connection
		WritableWorkbook wk = Workbook.createWorkbook(f);
		WritableSheet ws = wk.createSheet("Rodrigo", 0);
		for(int i=0 ; i<3 ; ++i) {		//loop for rows
			for(int j=0 ; j<3 ; ++j) {  //loop for columns
				Label L = new Label(j,i,"Hola");
				ws.addCell(L);
			}
		}
		wk.write();
		wk.close();
	}
	
	// Read excel file 'xls' with 'JXL' library.
	public void readData(String Path) throws BiffException, IOException {
		File f = new File(Path);
		Workbook wk = Workbook.getWorkbook(f);
		Sheet ws = wk.getSheet(0);
		int r = ws.getRows();
		int c = ws.getColumns();
		for(int i=0 ; i<r ; ++i) {			// loop for rows.
			
			for(int j=0 ; j<c ; ++j) {		// loop for columns.
				
				Cell c1 = ws.getCell(j, i);	// (col, row)
				System.out.println(c1.getContents());
			}	
		}
	}
	
	public static void main(String[] args) throws WriteException, IOException, BiffException {
		ExcelHandling e = new ExcelHandling();
		String path = "../Proyect1/Book1.xls";
		e.readData(path);
//		e.writeData(path);
	}	
}
