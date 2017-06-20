package cgi;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;

public class readExcel {
	readExcel(){
		 System.out.println("hey");
	}
public static void main() throws IOException{
	new readExcel();
	FileInputStream myStream = new FileInputStream("demo.xls");
	 NPOIFSFileSystem fs = new NPOIFSFileSystem(myStream);
	 HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
	 HSSFSheet sheet = wb.getSheetAt(0);
	 HSSFRow row;
	 HSSFCell cell;
	 int rows; // No of rows
	 rows = sheet.getPhysicalNumberOfRows();
	 int cols = 0; // No of columns
	 int tmp = 0;
	 for(int i=0;i<rows;i++){
		 row=sheet.getRow(i);
		 if(row!=null){
			 tmp = sheet.getRow(i).getPhysicalNumberOfCells();
			 if(tmp>cols)
				 cols = tmp;
			 for(int r = 0;r<cols;r++){
				 cell = row.getCell(r);
				 System.out.println(cell.getAddress());

			 }
		 }
	 }
}
}
